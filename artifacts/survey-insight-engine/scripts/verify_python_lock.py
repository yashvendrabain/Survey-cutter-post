#!/usr/bin/env python3
"""Verify the running Python environment exactly matches the root uv.lock.

The Survey Insight Engine's runtime packages live in the (gitignored)
``.pythonlibs/`` user-site, so the only auditable proof that the shipped
environment matches ``uv.lock`` is a reproducible check. Run this with the
project interpreter so ``importlib.metadata`` reflects the real environment::

    .pythonlibs/bin/python artifacts/survey-insight-engine/scripts/verify_python_lock.py

Exit code 0 means the environment matches the lock; non-zero lists every
discrepancy. The platform-resolved expectation comes from ``uv export`` so
that environment markers (e.g. colorama's ``sys_platform == 'win32'``) are
applied correctly for the current platform.
"""

from __future__ import annotations

import os
import re
import subprocess
import sys
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
    """Versions expected on *this* platform via `uv export`.

    `uv export` emits environment markers (e.g.
    ``colorama==0.4.6 ; sys_platform == 'win32'``). We evaluate each marker
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
    return {norm(d.metadata["Name"]): d.version for d in md.distributions() if d.metadata["Name"]}


def duplicate_dist_info() -> dict[str, list[str]]:
    dups: dict[str, list[str]] = {}
    seen: dict[str, list[str]] = {}
    for entry in sys.path:
        sp = Path(entry)
        if not sp.is_dir():
            continue
        for di in sp.glob("*.dist-info"):
            base = di.name[: -len(".dist-info")]
            m = re.match(r"^(.+)-([0-9][^-]*)$", base)
            if not m:
                continue
            seen.setdefault(norm(m.group(1)), []).append(di.name)
    for name, infos in seen.items():
        if len(set(infos)) > 1:
            dups[name] = sorted(set(infos))
    return dups


def main() -> int:
    root = repo_root()
    expected = lock_versions(root)
    installed = installed_versions()
    errors: list[str] = []

    # 1. Every locked package installed at the exact version.
    for name, ver in sorted(expected.items()):
        got = installed.get(name)
        if got is None:
            errors.append(f"MISSING: {name} (lock={ver})")
        elif got != ver:
            errors.append(f"VERSION MISMATCH: {name} installed={got} lock={ver}")

    # 2. No installed top-level package outside the lock (ignoring interpreter extras).
    for name, ver in sorted(installed.items()):
        if name not in expected and name not in ALLOWED_EXTRA:
            errors.append(f"UNEXPECTED (not in lock): {name}=={ver}")

    # 3. No duplicate/stale dist-info layers.
    for name, infos in sorted(duplicate_dist_info().items()):
        errors.append(f"DUPLICATE dist-info for {name}: {', '.join(infos)}")

    # 4. openai code/metadata coherence (the historical split-brain).
    try:
        import openai

        if openai.__version__ != md.version("openai"):
            errors.append(
                f"openai split-brain: __version__={openai.__version__} "
                f"metadata={md.version('openai')}"
            )
    except Exception as exc:  # pragma: no cover - openai always present here
        errors.append(f"openai import/version check failed: {exc}")

    if errors:
        print("FAIL: environment does not match uv.lock\n")
        for e in errors:
            print(f"  - {e}")
        return 1

    print(f"OK: {len(expected)} packages match uv.lock exactly")
    print(f"  openai coherent at {md.version('openai')}; no duplicate dist-info")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
