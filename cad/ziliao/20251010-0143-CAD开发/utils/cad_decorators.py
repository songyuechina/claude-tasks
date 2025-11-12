# -*- coding: utf-8 -*-
"""
Decorators and helpers to make CAD operations robust:
- Retry on transient COM errors (RPC busy/down)
- Wait for CAD to be idle before returning
- Optionally wait until a DWG document becomes visible in acad.Documents

Usage:

    from utils.cad_decorators import cad_operation, wait_cad_idle, wait_document_opened

    @cad_operation(retries=5, delay=1.0, wait_idle=True)
    def open_dwg(path: str):
        # your original logic...
        ...

    @cad_operation()
    def insert_and_explode_dwg(...):
        ...

You can also call wait_document_opened(acad, path) explicitly if the function
opens a DWG and you know the target file path.
"""
from __future__ import annotations

import time
from typing import Any, Callable, Optional


def _is_transient_com_error(e: Exception) -> bool:
    try:
        import pywintypes  # type: ignore
        if isinstance(e, pywintypes.com_error):
            code = getattr(e, "hresult", None)
            if code is None and e.args:
                code = e.args[0]
            RPC_BUSY = (-2147417846, -2147418111)  # server busy / call rejected
            RPC_DOWN = (-2147023174,)              # RPC server unavailable
            return code in (RPC_BUSY + RPC_DOWN)
    except Exception:
        pass
    # fallback: substring check
    s = str(e).lower()
    return ("server busy" in s) or ("rpc" in s and "unavailable" in s)


def wait_cad_idle(acad, *, min_quiet: float = 0.4, poll: float = 0.2, timeout: Optional[float] = None) -> bool:
    t0 = time.time()
    quiet_since: Optional[float] = None
    while True:
        try:
            st = acad.GetAcadState()
            if bool(getattr(st, "IsQuiescent", True)):
                if quiet_since is None:
                    quiet_since = time.time()
                if time.time() - quiet_since >= min_quiet:
                    return True
            else:
                quiet_since = None
        except Exception:
            # swallow and retry
            pass
        if timeout is not None and (time.time() - t0) > timeout:
            return False
        time.sleep(poll)


def _normpath(p: str) -> str:
    import os
    try:
        return os.path.normcase(os.path.abspath(p))
    except Exception:
        return p


def wait_document_opened(acad, path: str, *, poll: float = 0.25, timeout: Optional[float] = None) -> bool:
    want = _normpath(path)
    t0 = time.time()
    while True:
        try:
            for d in acad.Documents:
                try:
                    full = getattr(d, "FullName", None) or getattr(d, "Name", None)
                    if full and _normpath(full) == want:
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        if timeout is not None and (time.time() - t0) > timeout:
            return False
        time.sleep(poll)


def cad_operation(
    *,
    retries: int = 5,
    delay: float = 1.0,
    backoff: float = 1.0,
    wait_idle: bool = True,
    idle_quiet: float = 0.4,
    idle_timeout: Optional[float] = None,
    acad_getter: Optional[Callable[[], Any]] = None,
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Decorator to add retry + idle-wait semantics around CAD ops.

    - retries/delay/backoff: transient COM error handling
    - wait_idle: after func success, ensure CAD is idle before returning
    - acad_getter: supply a callable to get running acad; default uses Dispatch
    """
    def _decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
        import functools

        @functools.wraps(fn)
        def _wrapped(*args, **kwargs):
            last = None
            d = delay
            for attempt in range(1, max(1, retries) + 1):
                try:
                    res = fn(*args, **kwargs)
                    if wait_idle:
                        try:
                            if acad_getter is not None:
                                acad = acad_getter()
                            else:
                                import win32com.client  # type: ignore
                                acad = win32com.client.Dispatch("AutoCAD.Application")
                            wait_cad_idle(acad, min_quiet=idle_quiet, timeout=idle_timeout)
                        except Exception:
                            pass
                    return res
                except Exception as e:
                    last = e
                    if _is_transient_com_error(e) and attempt < retries:
                        try:
                            time.sleep(max(0.0, d))
                        except Exception:
                            pass
                        d *= backoff if backoff and backoff > 0 else 1.0
                        continue
                    raise

        return _wrapped

    return _decorator

def cad_sync_retry(
    *,
    max_attempts: int = 5,
    delay: float = 1.0,
    wait_idle: bool = True,
    idle_quiet: float = 0.4,
    idle_timeout: Optional[float] = 30.0,
) -> Callable[[Callable[..., Any]], Callable[..., Any]]:
    """Generic retry + CAD-idle synchronization decorator.

    - Retries on any exception raised by the wrapped function, waiting
      ``delay`` seconds between attempts, up to ``max_attempts``.
    - After a successful call, optionally waits until CAD becomes idle
      before returning, to ensure sequential command execution.

    The decorator attempts to locate a ``CadController`` instance from the
    function arguments (positional or keyword ``ctl``). If not found, it
    falls back to dispatching AutoCAD via COM for idle detection.
    """
    def _decorator(fn: Callable[..., Any]) -> Callable[..., Any]:
        import functools

        @functools.wraps(fn)
        def _wrapped(*args, **kwargs):
            last_exc: Any = None
            attempts = max(1, int(max_attempts))
            for i in range(1, attempts + 1):
                try:
                    result = fn(*args, **kwargs)
                    if wait_idle:
                        # Try to find controller
                        ctl = kwargs.get("ctl") if isinstance(kwargs, dict) else None
                        try:
                            from cad_automation import CadController  # type: ignore
                        except Exception:
                            CadController = None  # type: ignore
                        if ctl is None and CadController is not None:
                            for a in args:
                                try:
                                    if isinstance(a, CadController):
                                        ctl = a
                                        break
                                except Exception:
                                    pass
                        try:
                            if ctl is not None and hasattr(ctl, "wait_quiescent"):
                                ctl.wait_quiescent(timeout=idle_timeout or 30.0)
                            else:
                                try:
                                    import win32com.client  # type: ignore
                                    acad = win32com.client.Dispatch("AutoCAD.Application")
                                    wait_cad_idle(acad, min_quiet=idle_quiet, timeout=idle_timeout)
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    return result
                except Exception as e:
                    last_exc = e
                    if i >= attempts:
                        raise
                    try:
                        time.sleep(max(0.0, delay))
                    except Exception:
                        pass
            # If reached without return, re-raise the last exception
            if last_exc is not None:
                raise last_exc
            return None

        return _wrapped

    return _decorator
