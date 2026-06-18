#!/usr/bin/env python3
"""Realign the running Python environment so it exactly matches root uv.lock.

This is the *fixer* counterpart to ``verify_python_lock.py`` (the detector).
When ``uv.lock`` is bumped to newer versions, the runtime ``.pythonlibs/``
user-site keeps the stale packages (and leaves layered, stale ``*.dist-info``
dirs behind) until someone manually realigns it. Running this script brings
``.pythonlibs`` back in line with the lock automatically::

    .pythonlibs/bin/python artifacts/survey-insight-engine/scripts/realign_python_lock.py

It is idempotent and lightweight: when the environment already matches the lock
it does nothing. When it drifts, it:

1. Reads the platform-resolved expectation from ``uv export`` (environment
   markers evaluated, so e.g. win32-only ``colorama`` is skipped on Linux).
2. Removes packages installed but not in the lock (interpreter extras such as
   pip/setuptools/wheel are left alone), deleting their files via ``RECORD``.
3. For every package that drifts (wrong version, missing, or carrying
   duplicate/stale ``*.dist-info`` layers), wipes all of its existing
   ``*.dist-info`` layers (and their files via ``RECORD``) then force-reinstalls
   it at the exact locked version with ``pip --user --break-system-packages
   --no-deps --force-reinstall``.

After it runs, ``verify_python_lock.py`` passes with no manual steps.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import importlib.metadata as md

# Packages provided by the Nix interpreter itself, not part of uv.lock.
ALLOWED_EXTRA = {"pip", "setuptools", "wheel"}


def repo_root() -> Path:
    here = Path(__file__).resolve()
    for parent in here.parents:
        if (parent / "uv.lock").is_file():
            return parent
    raise SystemExit("Could not locate uv.lock above this script")


def norm(name: str) -> str:
    return name.lower().replace("_", "-")


def lock_versions(root: Path) -> dict[str, str]:
    """Versions expected on *this* platform via ``uv export``.

    ``uv export`` emits environment markers (e.g.
    ``colorama==0.4.6 ; sys_platform == 'win32'``). Each marker is evaluated
    against the current interpreter so platform-gated packages are only
    expected when they actually apply here.
    """
    from packaging.markers import Marker

    env = dict(os.environ)
    env.setdefault("UV_PYTHON", sys.executable)
    out = subprocess.run(
        [
            "uv",
            "export",
            "--all-groups",
            "--no-emit-project",
            "--format",
            "requirements-txt",
        ],
        cwd=root,
        env=env,
        capture_output=True,
        text=True,
    )
    if out.returncode != 0:
        raise SystemExit(f"`uv export` failed:\n{out.stderr}")
    versions: dict[str, str] = {}
    for raw in out.stdout.splitlines():
        line = raw.strip().rstrip("\\").strip()
        m = re.match(r"^([A-Za-z0-9._-]+)==([^ ;]+)(?:\s*;\s*(.+))?$", line)
        if not m:
            continue
        name, ver, marker = m.group(1), m.group(2), m.group(3)
        if marker and not Marker(marker).evaluate():
            continue
        versions[norm(name)] = ver
    return versions


def installed_versions() -> dict[str, str]:
    return {
        norm(d.metadata["Name"]): d.version
        for d in md.distributions()
        if d.metadata["Name"]
    }


def dist_info_dirs() -> dict[str, list[Path]]:
    """Map every package name to all of its ``*.dist-info`` dirs on sys.path."""
    result: dict[str, list[Path]] = {}
    for entry in sys.path:
        sp = Path(entry)
        if not sp.is_dir():
            continue
        for di in sp.glob("*.dist-info"):
            base = di.name[: -len(".dist-info")]
            m = re.match(r"^(.+)-([0-9][^-]*)$", base)
            if not m:
                continue
            result.setdefault(norm(m.group(1)), []).append(di)
    return result


def remove_dist_info(di: Path) -> None:
    """Delete a ``*.dist-info`` dir and the files it records."""
    site = di.parent
    record = di / "RECORD"
    if record.is_file():
        for line in record.read_text(encoding="utf-8", errors="replace").splitlines():
            rel = line.split(",", 1)[0].strip()
            if not rel or rel.endswith("/"):
                continue
            target = (site / rel)
            try:
                resolved = target.resolve()
            except OSError:
                continue
            # Stay within the site directory to avoid deleting anything odd.
            try:
                resolved.relative_to(site.resolve())
            except ValueError:
                continue
            try:
                if resolved.is_file() or resolved.is_symlink():
                    resolved.unlink()
            except OSError:
                pass
    shutil.rmtree(di, ignore_errors=True)


def pip_install(reqs: dict[str, str]) -> None:
    """Force-reinstall the given ``name -> version`` packages at exact version."""
    with tempfile.NamedTemporaryFile(
        "w", suffix=".txt", delete=False, encoding="utf-8"
    ) as fh:
        for name, ver in sorted(reqs.items()):
            fh.write(f"{name}=={ver}\n")
        req_path = fh.name
    try:
        cmd = [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--user",
            "--break-system-packages",
            "--no-deps",
            "--force-reinstall",
            "-r",
            req_path,
        ]
        print(f"$ {' '.join(cmd)}")
        proc = subprocess.run(cmd)
        if proc.returncode != 0:
            raise SystemExit(f"pip install failed (exit {proc.returncode})")
    finally:
        try:
            os.unlink(req_path)
        except OSError:
            pass


def main() -> int:
    root = repo_root()
    expected = lock_versions(root)
    installed = installed_versions()
    dist_infos = dist_info_dirs()

    # 1. Remove anything installed that is not in the lock (keep interpreter extras).
    for name in sorted(installed):
        if name in expected or name in ALLOWED_EXTRA:
            continue
        print(f"removing not-in-lock package: {name}=={installed[name]}")
        for di in dist_infos.get(name, []):
            remove_dist_info(di)

    # 2. Decide what to reinstall: wrong/missing version, or duplicate dist-info.
    to_reinstall: dict[str, str] = {}
    for name, ver in expected.items():
        got = installed.get(name)
        dups = dist_infos.get(name, [])
        if got != ver or len({d.name for d in dups}) > 1:
            to_reinstall[name] = ver

    if not to_reinstall:
        print(f"OK: {len(expected)} packages already match uv.lock; nothing to do")
        return 0

    # 3. Wipe every existing dist-info layer for the drifted packages so the
    #    force-reinstall leaves exactly one coherent version behind.
    for name in sorted(to_reinstall):
        for di in dist_infos.get(name, []):
            print(f"clearing stale dist-info: {di.name}")
            remove_dist_info(di)

    print(f"reinstalling {len(to_reinstall)} package(s) to match uv.lock")
    pip_install(to_reinstall)

    print(f"done: realigned {len(to_reinstall)} package(s) to uv.lock")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
