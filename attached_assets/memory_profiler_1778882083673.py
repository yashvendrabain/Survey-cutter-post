"""Low-overhead memory profiling helpers for workbook exports."""

from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
import os
import sys
import tracemalloc
from typing import Iterator

try:
    import psutil
except ModuleNotFoundError:  # pragma: no cover - requirements installs psutil.
    psutil = None


@dataclass(frozen=True)
class _StepRecord:
    label: str
    rss_start: int | None
    rss_end: int | None
    rss_delta: int | None
    tracemalloc_peak: int


_STEP_LOG: list[_StepRecord] = []
_PROFILING_ENABLED = False
_PROCESS = None


@contextmanager
def memory_step(label: str) -> Iterator[None]:
    """Capture RSS and tracemalloc peak for a named export step."""

    if not _PROFILING_ENABLED:
        yield
        return

    if not tracemalloc.is_tracing():
        tracemalloc.start()

    rss_start = _rss_bytes()
    if hasattr(tracemalloc, "reset_peak"):
        tracemalloc.reset_peak()
    try:
        yield
    finally:
        _current, peak = tracemalloc.get_traced_memory()
        rss_end = _rss_bytes()
        rss_delta = (
            None
            if rss_start is None or rss_end is None
            else rss_end - rss_start
        )
        _STEP_LOG.append(
            _StepRecord(
                label=label,
                rss_start=rss_start,
                rss_end=rss_end,
                rss_delta=rss_delta,
                tracemalloc_peak=int(peak),
            )
        )


def get_report() -> str:
    """Return a formatted memory profile table."""

    header = "step | rss_start | rss_end | rss_delta | tracemalloc_peak"
    separator = "--- | ---: | ---: | ---: | ---:"
    rows = [
        " | ".join(
            (
                record.label,
                _format_bytes(record.rss_start),
                _format_bytes(record.rss_end),
                _format_bytes(record.rss_delta),
                _format_bytes(record.tracemalloc_peak),
            )
        )
        for record in _STEP_LOG
    ]
    return "\n".join([header, separator, *rows])


def reset_log() -> None:
    """Clear captured step records between runs."""

    _STEP_LOG.clear()


def enable_profiling() -> None:
    """Enable memory profiling and start tracemalloc."""

    global _PROFILING_ENABLED, _PROCESS
    if psutil is not None and _PROCESS is None:
        _PROCESS = psutil.Process(os.getpid())
    if not tracemalloc.is_tracing():
        tracemalloc.start()
    _PROFILING_ENABLED = True


def disable_profiling() -> None:
    """Disable memory profiling without clearing the last report."""

    global _PROFILING_ENABLED
    _PROFILING_ENABLED = False


def is_profiling_enabled() -> bool:
    return _PROFILING_ENABLED


def _rss_bytes() -> int | None:
    if psutil is not None:
        global _PROCESS
        if _PROCESS is None:
            _PROCESS = psutil.Process(os.getpid())
        return int(_PROCESS.memory_info().rss)

    if sys.platform != "win32":
        try:
            import resource

            usage = resource.getrusage(resource.RUSAGE_SELF)
            rss = int(usage.ru_maxrss)
            return rss if sys.platform == "darwin" else rss * 1024
        except Exception:
            return None
    return None


def _format_bytes(value: int | None) -> str:
    if value is None:
        return "n/a"
    sign = "-" if value < 0 else ""
    absolute = abs(float(value))
    return f"{sign}{absolute / (1024 * 1024):,.1f} MiB"
