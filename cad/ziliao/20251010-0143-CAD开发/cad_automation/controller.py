# -*- coding: utf-8 -*-
from __future__ import annotations

import importlib.util
import json
import os
import sys
import time
from pathlib import Path
from types import ModuleType
from typing import Iterable, Optional, Sequence, Tuple


DEFAULT_CAD_OPS_PATH = r"D:/Myprogramsystem/cad/CAD基本操作.py"


def _ensure_logs_dir(logs_dir: Path) -> None:
    logs_dir.mkdir(parents=True, exist_ok=True)


def load_cad_ops_module(path: Optional[str] = None) -> ModuleType:
    """Load external CAD operations module from file path.
    Resolution order: explicit arg -> env CAD_OPS_PATH -> config external_module_path -> DEFAULT_CAD_OPS_PATH
    """
    module_path: Optional[str] = path
    if not module_path:
        module_path = os.environ.get("CAD_OPS_PATH")
    if not module_path:
        cfg_path = Path(__file__).with_name("config.json")
        if cfg_path.exists():
            try:
                data = json.loads(cfg_path.read_text(encoding="utf-8"))
                module_path = data.get("external_module_path")
            except Exception:
                module_path = None
    if not module_path:
        module_path = DEFAULT_CAD_OPS_PATH

    module_path = str(module_path)
    if not os.path.isfile(module_path):
        raise FileNotFoundError(f"外部脚本未找到: {module_path}")

    os.environ.setdefault("PYTHONUTF8", "1")

    spec = importlib.util.spec_from_file_location("cad_ops", module_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {module_path}")
    mod = importlib.util.module_from_spec(spec)
    mod_dir = os.path.dirname(module_path)
    if mod_dir and mod_dir not in sys.path:
        sys.path.insert(0, mod_dir)
    sys.modules["cad_ops"] = mod
    spec.loader.exec_module(mod)  # type: ignore[attr-defined]
    return mod


class CadController:
    """Thin wrapper around external CAD ops providing a stable interface."""

    def __init__(
        self,
        cad_ops_path: Optional[str] = None,
        logs_dir: str | os.PathLike = "artifacts/logs",
        tarch_root: Optional[str] = None,
    ) -> None:
        self.logs_dir = Path(logs_dir)
        _ensure_logs_dir(self.logs_dir)
        self.log_file = self.logs_dir / "cad_automation.log"
        self._locks_dir = Path("artifacts/locks")
        self._locks_dir.mkdir(parents=True, exist_ok=True)

        self.ops = load_cad_ops_module(cad_ops_path)
        self.tarch_root = tarch_root or r"C:\\Tangent\\TArchT20V9"

        self._proc = None
        self._opening_paths: set[str] = set()
        # Opening wait tolerance (seconds) per requirement
        self.open_timeout_s: int = 120

    # --------------- utilities ---------------
    def _log(self, msg: str) -> None:
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}\n"
        try:
            if self.log_file.exists():
                prev = self.log_file.read_text(encoding="utf-8")
                self.log_file.write_text(prev + line, encoding="utf-8")
            else:
                self.log_file.write_text(line, encoding="utf-8")
        except Exception:
            pass

    def _ensure_li(self) -> None:
        if hasattr(self.ops, "li_new"):
            self.ops.li_new()
        elif hasattr(self.ops, "li"):
            self.ops.li()
        else:
            self.ops.get_acad_doc()

    def _cad_process_count(self) -> int:
        try:
            if hasattr(self.ops, "jingchengshu_wenjian"):
                return int(self.ops.jingchengshu_wenjian())
        except Exception:
            pass
        try:
            self.ops.get_acad_doc()
            return 1
        except Exception:
            return 0

    # ----------- TArch detection helpers -----------
    def _psutil_names(self) -> list[str]:
        try:
            import psutil  # type: ignore
        except Exception:
            return []
        names: list[str] = []
        try:
            for p in psutil.process_iter(["name"]):
                n = (p.info.get("name") or "").strip().lower()
                if n:
                    names.append(n)
        except Exception:
            pass
        return names

    def is_tarch_running(self) -> bool:
        names = self._psutil_names()
        keys = ("tarcht20v9.exe", "tarch.exe", "tarcht20.exe")
        return any(k in names for k in keys)

    def _ensure_one_process_before_open(self, wait_s: float = 0.3, *, strict: bool = False) -> None:
        """Ensure a sane process state before open.

        Default (strict=False):
        - cnt == 0  -> start_tarch_v9()
        - cnt  > 1  -> ensure_single_process()（尽量收敛到单进程，不强制重启）
        - cnt == 1  -> 仅确认空闲

        Strict mode (strict=True):
        - cnt != 1  -> 调用 single_unsaved_state()，用于强约束场景（如双文件确定状态前的准备）
        """
        try:
            cnt = int(self._cad_process_count())
        except Exception:
            cnt = 0
        try:
            if strict and cnt != 1:
                # 强约束：直接收敛到单未保存
                self.single_unsaved_state()
                return
            if cnt == 0:
                self.start_tarch_v9()
            elif cnt > 1:
                self.ensure_single_process()
            # 短暂等待稳定
            self.wait_quiescent(timeout=20)
            import time as _time
            _time.sleep(max(0.1, min(3.0, float(wait_s))))
        except Exception:
            pass

    def cad_process_count(self, print_it: bool = True) -> int:
        """Public helper: return current acad.exe process count; optionally print it.

        Printing format: "[cad-process-count] <n>" for easy grepping in logs.
        """
        try:
            cnt = int(self._cad_process_count())
        except Exception:
            cnt = 0
        if print_it:
            try:
                print(f"[cad-process-count] {cnt}")
            except Exception:
                pass
        return cnt

    def ensure_single_process(self) -> None:
        """Ensure there is at most one acad.exe process running."""
        try:
            cnt = self._cad_process_count()
            if cnt <= 1:
                return
            if hasattr(self.ops, "close_extra_acad_processes"):
                try:
                    self.ops.close_extra_acad_processes()
                    return
                except Exception:
                    pass
            # Fallback: close all and let caller restart
            self.close_all_cad()
        except Exception:
            pass

    # --------------- app level ---------------
    def start_tarch_v9(self, root: Optional[str] = None, max_retries: int = 3) -> tuple[str, str]:
        PTH = root or self.tarch_root
        self._log(f"start_tarch_v9: PTH={PTH}")
        try:
            # shrink to <=1
            self.ensure_single_process()
            if self._cad_process_count() >= 1:
                self._log("reuse existing acad.exe")
                self._ensure_li()
                return True
            self._proc = self.ops.start_applicationV9(PTH=PTH, max_retries=max_retries)
            # wait COM ready
            t0 = time.time()
            while time.time() - t0 < 60:
                try:
                    self.ops.get_acad_doc()
                    break
                except Exception:
                    time.sleep(0.5)
            self._ensure_li()
            return self._cad_process_count() >= 1
        except Exception as e:
            self._log(f"start_tarch_v9 failed: {e!r}")
            return False

    # --------------- waits ---------------
    def wait_quiescent(self, timeout: float = 60.0, poll: float = 0.5) -> tuple[str, str]:
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                app, _ = self.ops.get_acad_doc()
                state = app.GetAcadState()
                if bool(getattr(state, "IsQuiescent", True)):
                    return True
            except Exception:
                pass
            time.sleep(poll)
        return False

    def _normalize_path(self, p: str) -> str:
        return os.path.normcase(os.path.abspath(p))

    def _is_doc_open(self, path: str) -> tuple[str, str]:
        want = self._normalize_path(path)
        try:
            for d in self.ops.acad.Documents:
                try:
                    full = getattr(d, "FullName", None) or getattr(d, "Name", None)
                    if full and self._normalize_path(full) == want:
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    def _is_doc_open_by_name(self, name: str) -> bool:
        try:
            for d in self.ops.acad.Documents:
                try:
                    nm = str(getattr(d, "Name", ""))
                    if nm and nm.lower() == str(name).lower():
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    def _activate_document(self, path: str) -> tuple[str, str]:
        want = self._normalize_path(path)
        try:
            for d in self.ops.acad.Documents:
                try:
                    full = getattr(d, "FullName", None) or getattr(d, "Name", None)
                    if full and self._normalize_path(full) == want:
                        d.Activate()
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    def wait_document_opened(self, path: str, timeout: float = 45.0, poll: float = 0.5) -> tuple[str, str]:
        t0 = time.time()
        while time.time() - t0 < timeout:
            if self._is_doc_open(path):
                return True
            time.sleep(poll)
        return False

    def wait_document_closed(self, path: str, timeout: float = 45.0, poll: float = 0.5) -> tuple[str, str]:
        t0 = time.time()
        while time.time() - t0 < timeout:
            if not self._is_doc_open(path):
                return True
            time.sleep(poll)
        return False

    # --------------- window helpers ---------------
    def _minimize_other_windows(self, keep_keys: Sequence[str] = ("AutoCAD", "TArch", "天正")) -> None:
        try:
            import pygetwindow as gw  # type: ignore
        except Exception:
            return
        try:
            alls = gw.getAllTitles()
            for t in alls:
                if not t.strip():
                    continue
                low = t.lower()
                if any(k.lower() in low for k in keep_keys):
                    continue
                try:
                    w = gw.getWindowsWithTitle(t)
                    for win in w:
                        try:
                            win.minimize()
                        except Exception:
                            pass
                except Exception:
                    pass
        except Exception:
            pass

    def activate_cad_window(self, title_key: str = "AutoCAD", click_titlebar: bool = True, minimize_others: bool = True) -> tuple[str, str]:
        try:
            if minimize_others:
                self._minimize_other_windows()
            if hasattr(self.ops, "activate_window_by_title"):
                return bool(self.ops.activate_window_by_title(title_key, click_titlebar))
        except Exception as e:
            self._log(f"activate_cad_window error: {e!r}")
        return False

    def _close_error_windows(self, titles: Optional[Sequence[str]] = None, repeat: int = 3) -> None:
        """Best-effort close for known AutoCAD error-report windows by title.

        Uses external ops helpers when available:
        - activate_window_by_title(title, click_titlebar=True) to bring it front and get size
        - click_in_window(title, offset_x, offset_y, click_titlebar=True) to click the close button

        This is defensive and silent-on-error, so it won't break main flows.
        """
        titles = titles or (
            "AutoCAD 错误报告",
            "错误报告",
            "AutoCAD 异常终止",
            "AutoCAD Error Report",
        )
        for _ in range(max(1, repeat)):
            closed_any = False
            for title in titles:
                try:
                    box = None
                    if hasattr(self.ops, "activate_window_by_title"):
                        try:
                            box = self.ops.activate_window_by_title(str(title), True)
                        except Exception:
                            box = None
                    # Try precise close button click if we know the size
                    if hasattr(self.ops, "click_in_window"):
                        try:
                            if isinstance(box, (tuple, list)) and len(box) >= 4:
                                _, _, w, _ = box
                                x = max(10, int(w) - 10)
                            else:
                                x = 240
                            self.ops.click_in_window(str(title), x, 10, click_titlebar=True)
                            closed_any = True
                            time.sleep(0.2)
                            continue
                        except Exception:
                            pass
                    # Fallback: send Alt+F4 to foreground
                    try:
                        import pyautogui  # type: ignore
                        pyautogui.hotkey("alt", "f4")
                        closed_any = True
                        time.sleep(0.2)
                    except Exception:
                        pass
                except Exception:
                    continue
            if not closed_any:
                break
        try:
            self.wait_quiescent(timeout=10)
        except Exception:
            pass

    # --------------- process ---------------
    def close_all_cad(self) -> None:
        self._log("close_all_cad")
        try:
            self.ops.close_all_cad_processes()
        finally:
            self._proc = None

    def close_all_cad_and_wait(self, timeout: float = 60.0, interval: float = 0.5) -> bool:
        """Close all CAD processes via external ops and wait until the count is 0.

        只调用外部 `close_all_cad_processes()`（来自 CAD 基本操作.py），在超时时间内轮询直至进程数为 0。
        返回 True 表示已降到 0；否则 False。
        """
        t0 = time.time()
        # 首次请求关闭
        try:
            self.ops.close_all_cad_processes()
        except Exception:
            pass
        while time.time() - t0 < timeout:
            try:
                cnt = int(self._cad_process_count())
            except Exception:
                cnt = 0
            if cnt <= 0:
                try:
                    print(f"[proc-count-zero] {cnt}")
                except Exception:
                    pass
                return True
            # 继续请求关闭并稍作等待
            try:
                self.ops.close_all_cad_processes()
            except Exception:
                pass
            time.sleep(interval)
        return False

    # --------------- document ops ---------------
    def open_dwg(self, path: str) -> tuple[str, str]:
        """Open a DWG by absolute path. If already open, just activate.
        Uses self.open_timeout_s (default 120s) for waits.
        Important: to avoid opening multiple read-only copies of the same file,
        this function issues at most ONE Documents.Open() call per request and
        then waits until the document is visible in acad.Documents.
        """
        # Ensure sane process state before opening (non-destructive)
        self._ensure_one_process_before_open(wait_s=0.3, strict=False)
        path = self._normalize_path(path)

        # Win32 short path for COM robustness
        def _short_path(p: str) -> str:
            try:
                import ctypes
                from ctypes import wintypes
                GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
                GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
                GetShortPathNameW.restype = wintypes.DWORD
                buf = ctypes.create_unicode_buffer(260)
                ret = GetShortPathNameW(p, buf, len(buf))
                return buf.value if ret else p
            except Exception:
                return p

        spath = _short_path(path)
        self._log(f"open_dwg: {path}")
        self._ensure_li()
        self.wait_quiescent(timeout=self.open_timeout_s)

        if path in self._opening_paths:
            self._log(f"skip open: already opening {path}")
            return self.wait_document_opened(path, timeout=self.open_timeout_s)

        if self._is_doc_open(path):
            self._activate_document(path)
            return True

        # file lock to avoid concurrent open to the same path
        from contextlib import contextmanager
        @contextmanager
        def _path_lock(p: str):
            import hashlib
            lock_dir = self._locks_dir
            h = hashlib.sha1(self._normalize_path(p).encode("utf-8")).hexdigest()
            fp = lock_dir / f"open_{h}.lock"
            for _ in range(3):
                try:
                    fh = os.open(str(fp), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(fh)
                    break
                except Exception:
                    time.sleep(0.2)
            try:
                yield
            finally:
                try:
                    fp.unlink(missing_ok=True)  # type: ignore[arg-type]
                except Exception:
                    pass

        with _path_lock(path):
            if self._is_doc_open(path):
                self._activate_document(path)
                return True
            try:
                self._opening_paths.add(path)
                import time as _time
                ok = False
                # Issue at most one Open() call and then wait long enough
                already_opening = False
                try:
                    # Final pre-check
                    if not self._is_doc_open(path):
                        self.ops.acad.Documents.Open(spath)
                        already_opening = True
                except Exception:
                    # If CAD is busy, give it a brief moment and continue to wait
                    already_opening = True
                    _time.sleep(0.5)

                # Wait until the document becomes visible (up to open_timeout_s)
                base = os.path.basename(path)
                if self.wait_document_opened(path, timeout=self.open_timeout_s):
                    ok = True
                else:
                    # Fallback: check by basename to be robust with non-ASCII paths
                    t0 = _time.time()
                    while _time.time() - t0 < 10.0:
                        if self._is_doc_open_by_name(base):
                            ok = True
                            break
                        _time.sleep(0.3)
                    if not ok:
                        # As a last resort if present by path
                        ok = bool(self._is_doc_open(path))

                self.wait_quiescent(timeout=self.open_timeout_s)
                if ok:
                    self._activate_document(path)
                return ok
            except Exception as e:
                self._log(f"open_dwg error: {e!r}")
                return False
            finally:
                self._opening_paths.discard(path)
    def open_dwgs(self, paths: Sequence[str]) -> int:
        """Strictly open DWGs one-by-one with cooperative sequencing.

        Rules:
        - For each target path, check if it's already open by full path; if yes, skip.
        - Also check if a document with the same basename is already open; if yes, skip
          to avoid redundant read-only copies of the same file.
        - Call ``open_dwg`` at most once per file, then wait until it appears in
          acad.Documents and CAD becomes quiescent before proceeding to the next file.
        - Inserts a small grace delay (0.3s) between consecutive opens.
        """
        # Ensure sane process state before opening (non-destructive)
        self._ensure_one_process_before_open(wait_s=0.3, strict=False)
        ok = 0
        names_cache: list[str] | None = None

        for raw in paths:
            if not raw:
                continue
            p = self._normalize_path(str(raw))

            # Refresh current names lazily
            try:
                names_cache = self.list_open_documents()
            except Exception:
                names_cache = names_cache or []

            # Path-level idempotency
            if self._is_doc_open(p):
                continue

            # Basename-level guard to avoid duplicate read-only copies
            base = os.path.basename(p)
            if names_cache and base in names_cache:
                # Already a document with the same name; skip to avoid duplicates
                continue

            # Ensure CAD is calm before opening next
            self.wait_quiescent(timeout=min(self.open_timeout_s, 30))

            if self.open_dwg(p):
                ok += 1
                # Update names cache after success and leave a small grace delay
                try:
                    names_cache = self.list_open_documents()
                except Exception:
                    names_cache = None
                try:
                    import time as _time
                    _time.sleep(0.3)
                except Exception:
                    pass
        return ok

    def list_open_documents(self) -> list[str]:
        self._ensure_li()
        try:
            if hasattr(self.ops, "get_open_document_names"):
                return list(self.ops.get_open_document_names())
            return [doc.Name for doc in self.ops.acad.Documents]
        except Exception as e:
            self._log(f"list_open_documents error: {e!r}")
            return []

    def doc_count(self) -> int:
        self._ensure_li()
        try:
            if hasattr(self.ops, "dwgs_count"):
                return int(self.ops.dwgs_count())
            return int(self.ops.acad.Documents.Count)
        except Exception:
            return 0

    def close_current(self) -> None:
        self._ensure_li()
        try:
            try:
                cur_path = getattr(self.ops.doc, "FullName", None)
            except Exception:
                cur_path = None
            self.ops.doc.Close(False)
            if cur_path:
                self.wait_document_closed(cur_path, timeout=60)
            self.wait_quiescent(timeout=30)
        except Exception:
            pass

    def close_by_name(self, name: str) -> None:
        self._ensure_li()
        for d in list(self.ops.acad.Documents):
            try:
                if str(getattr(d, "Name", "")).lower() == name.lower():
                    full = getattr(d, "FullName", None)
                    d.Close(False)
                    if full:
                        self.wait_document_closed(full, timeout=60)
                    self.wait_quiescent(timeout=30)
                    break
            except Exception:
                pass

    def save(self) -> None:
        self._ensure_li()
        try:
            self.wait_quiescent(timeout=30)
            try:
                self.ops.doc.Save()
                return
            except Exception:
                full = getattr(self.ops.doc, "FullName", None)
                if full:
                    self.save_as(full)
                else:
                    raise
        except Exception as e:
            self._log(f"save error: {e!r}")

    def save_as(self, out_path: str) -> None:
        self._ensure_li()
        def _short_path(p: str) -> str:
            try:
                import ctypes
                from ctypes import wintypes
                GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
                GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
                GetShortPathNameW.restype = wintypes.DWORD
                buf = ctypes.create_unicode_buffer(260)
                ret = GetShortPathNameW(p, buf, len(buf))
                return buf.value if ret else p
            except Exception:
                return p
        try:
            self.wait_quiescent(timeout=30)
            sp = _short_path(os.path.abspath(out_path))
            self.ops.doc.SaveAs(sp)
            self.wait_quiescent(timeout=30)
        except Exception as e:
            self._log(f"save_as error: {e!r}")

    # --------------- examples ---------------
    def draw_line(self, p1: Tuple[float, float, float], p2: Tuple[float, float, float]):
        self._ensure_li()
        pt1 = self.ops.vtpnt(*p1)
        pt2 = self.ops.vtpnt(*p2)
        line = self.ops.mp.AddLine(pt1, pt2)
        return line

    # --------------- convenience ---------------
    def start_and_open(self, dwgs: Sequence[str], tarch_root: Optional[str] = None) -> int:
        ok = self.start_tarch_v9(root=tarch_root)
        if not ok:
            return 0
        return self.open_dwgs(dwgs)

    def restart_fresh_tarch_v9(self, root: Optional[str] = None, *, max_retries: int = 3, retry_delay: float = 2.0) -> bool:
        """Close all CAD processes then start a fresh TArch V9 + CAD2021 instance.

        This explicitly calls the external ops function ``start_applicationV9`` with
        the signature ``(PTH, max_retries, retry_delay)`` when available, and
        falls back to ``(PTH, max_retries)`` for older versions.

        Returns True if a process is up and COM is ready.
        """
        try:
            # Close everything first to guarantee a brand-new process
            try:
                self.close_all_cad()
            except Exception:
                pass

            PTH = root or self.tarch_root
            # Invoke external starter with retry_delay if supported
            try:
                self.ops.start_applicationV9(PTH=PTH, max_retries=max_retries, retry_delay=retry_delay)
            except TypeError:
                self.ops.start_applicationV9(PTH=PTH, max_retries=max_retries)

            # wait COM ready
            t0 = time.time()
            while time.time() - t0 < 60:
                try:
                    self.ops.get_acad_doc()
                    break
                except Exception:
                    time.sleep(0.5)
            self._ensure_li()
            return self._cad_process_count() >= 1
        except Exception as e:
            self._log(f"restart_fresh_tarch_v9 failed: {e!r}")
            return False

    def single_unsaved_state(self, wait_s: float = 2.0) -> str:
        """Ensure a "single-file indeterminate" state.

        What it guarantees
        - Single CAD process, single DWG in the UI.
        - Does not guarantee which DWG it is; may create a blank unsaved DWG when none exists.

        Rules
        - Close all running CAD processes, then call ``start_applicationV9`` to start
          TArch V9 + CAD2021. Note: This usually opens a default unsaved DWG; when
          you subsequently open a specific file, that default document will close
          automatically by CAD – this property is intentional and relied upon here.
        - If for any reason multiple DWGs appear, keep the currently active DWG and
          close others without saving until only one remains.
        - Already one DWG: keep as-is.

        Returns
        - The active DWG basename (e.g. "Drawing1.dwg"), or "" on failure.

        Usage
        - From Python: ``CadController().single_unsaved_state()``
        - From CLI (equivalent effect): ``python cad_cli.py new`` (no path argument)
        """
        # 1) Process-aware restart sequence
        try:
            cnt = self._cad_process_count()
        except Exception:
            cnt = 0

        # If there are running CAD processes, close them all first
        if cnt > 0:
            try:
                print(f"[proc-count-before] {cnt}")
            except Exception:
                pass
            self.close_all_cad_and_wait(timeout=90.0)

        # Start TArch V9 + CAD2021 via external start_applicationV9
        try:
            PTH = self.tarch_root
            try:
                self.ops.start_applicationV9(PTH=PTH, max_retries=3, retry_delay=2.0)
            except TypeError:
                self.ops.start_applicationV9(PTH=PTH, max_retries=3)
        except Exception as e:
            self._log(f"single_unsaved_state start_applicationV9 error: {e!r}")
            return ""

        # Wait COM ready
        t0 = time.time()
        while time.time() - t0 < 60:
            try:
                self.ops.get_acad_doc()
                break
            except Exception:
                time.sleep(0.5)
        self._ensure_li()
        # Verify TArch presence; if not detected, attempt one more start
        try:
            if not self.is_tarch_running():
                try:
                    self.ops.start_applicationV9(PTH=self.tarch_root, max_retries=3, retry_delay=2.0)
                except TypeError:
                    self.ops.start_applicationV9(PTH=self.tarch_root, max_retries=3)
                time.sleep(1.0)
        except Exception:
            pass
        try:
            print(f"[proc-count-after-start] {self._cad_process_count()}")
        except Exception:
            pass

        # 2b) Optionally shrink to a single acad.exe process if multiple exist
        try:
            self.ensure_single_process()
            try:
                print(f"[proc-count-after-ensure] {self._cad_process_count()}")
            except Exception:
                pass
        except Exception:
            pass

        # 2c) Close possible error-report windows if they popped up
        self._close_error_windows()

        # 3) Determine current documents
        try:
            docs = list(self.ops.acad.Documents)
        except Exception:
            docs = []

        # 4) If none, create a new unsaved drawing
        if len(docs) == 0:
            try:
                self.ops.acad.Documents.Add()
                self.wait_quiescent(timeout=30)
            except Exception:
                try:
                    self.ops.get_acad_doc()
                    self.ops.acad.Documents.Add()
                    self.wait_quiescent(timeout=30)
                except Exception:
                    return ""
        # 5) If more than one, close others but keep active (robust loop until <=1)
        else:
            try:
                for _ in range(15):
                    try:
                        active = getattr(self.ops.acad, "ActiveDocument", None)
                    except Exception:
                        active = None
                    cnt = 0
                    try:
                        cnt = int(getattr(self.ops.acad.Documents, "Count", 0))
                    except Exception:
                        cnt = 0
                    if cnt <= 1:
                        break
                    removed = False
                    for d in list(self.ops.acad.Documents):
                        try:
                            if active is not None and d is active:
                                continue
                            full = getattr(d, "FullName", None)
                            d.Close(False)
                            if full:
                                self.wait_document_closed(full, timeout=60)
                            removed = True
                            break
                        except Exception:
                            pass
                    self.wait_quiescent(timeout=20)
                    if not removed:
                        break
            except Exception:
                pass
        # 6) One more sweep for error-report windows
        self._close_error_windows()

        # 5) Bring window forward and return active name
        try:
            self.activate_cad_window(title_key="AutoCAD", click_titlebar=True, minimize_others=True)
        except Exception:
            pass
        try:
            name = str(getattr(getattr(self.ops.acad, "ActiveDocument", None), "Name", ""))
        except Exception:
            name = ""
        # Per requirement, leave a grace window between 1–3 seconds
        try:
            _ws = float(wait_s)
        except Exception:
            _ws = 2.0
        _ws = max(1.0, min(3.0, _ws))
        try:
            print(f"[post-state-wait] { _ws:.1f}s")
        except Exception:
            pass
        time.sleep(_ws)
        if name:
            print(f"[single-indeterminate] {name}")
        return name

    def standardize_state(self, target_doc_path: Optional[str] = None, wait_s: float = 1.0) -> str:
        """Ensure a "single-file definite" (stable) state.

        What it guarantees
        - Single CAD process, exactly one DWG remains and is active.
        - The remaining DWG is deterministic:
          - If ``target_doc_path`` is provided, it will be opened and kept as the sole document.
          - If ``target_doc_path`` is None, it keeps the currently active DWG and closes the rest.

        Returns
        - The active DWG basename (e.g. "PlanA.dwg") on success, or "" on failure.

        Usage
        - From Python: ``CadController().standardize_state(<optional_dwg_path>)``
        - From CLI: ``python cad_cli.py standard <optional_dwg_path>``
        """
        # single process
        self.ensure_single_process()
        if not self.start_tarch_v9():
            return False
        self._ensure_li()
        if target_doc_path:
            self.open_dwg(target_doc_path)
            self._activate_document(target_doc_path)
        else:
            self.ops.get_acad_doc()
        # Close any potential error-report popups before and after closing others
        self._close_error_windows()
        try:
            if hasattr(self.ops, "close_all_except_active_safe"):
                try:
                    self.ops.close_all_except_active_safe()
                except Exception as e:
                    self._log(f"close_all_except_active_safe error: {e!r}")
                    for _ in range(10):
                        self.wait_quiescent(timeout=15)
                        active = getattr(self.ops.acad, "ActiveDocument", None)
                        cnt = int(getattr(self.ops.acad.Documents, "Count", 0))
                        if cnt <= 1:
                            break
                        for d in list(self.ops.acad.Documents):
                            try:
                                if active is not None and d is active:
                                    continue
                                name = str(getattr(d, "Name", ""))
                                if hasattr(self.ops, "close_dwg_by_name") and name:
                                    self.ops.close_dwg_by_name(name)
                                else:
                                    full = getattr(d, "FullName", None)
                                    d.Close(False)
                                    if full:
                                        self.wait_document_closed(full, timeout=60)
                                break
                            except Exception:
                                pass
            else:
                for _ in range(10):
                    self.wait_quiescent(timeout=15)
                    active = getattr(self.ops.acad, "ActiveDocument", None)
                    cnt = int(getattr(self.ops.acad.Documents, "Count", 0))
                    if cnt <= 1:
                        break
                    for d in list(self.ops.acad.Documents):
                        try:
                            if active is not None and d is active:
                                continue
                            name = str(getattr(d, "Name", ""))
                            if hasattr(self.ops, "close_dwg_by_name") and name:
                                self.ops.close_dwg_by_name(name)
                            else:
                                full = getattr(d, "FullName", None)
                                d.Close(False)
                                if full:
                                    self.wait_document_closed(full, timeout=60)
                            break
                        except Exception:
                            pass
        finally:
            self.wait_quiescent(timeout=45)
            time.sleep(wait_s)

        self.activate_cad_window(title_key="AutoCAD", click_titlebar=True, minimize_others=True)
        self._close_error_windows()
        # collect current single document name and return it
        try:
            names = [str(getattr(d, "Name", "")) for d in self.ops.acad.Documents]
        except Exception:
            names = []
        base = names[0] if names else ""
        if base:
            print(f"[single] {base}")
        return base


    def standardize_two_documents(self, doc_a: str, doc_b: str, *, active: Optional[str] = None, wait_s: float = 1.0) -> tuple[str, str]:
        """Ensure a "two-file definite" state in a single CAD process.

        What it guarantees
        - Single CAD process, exactly two specified DWGs are opened and remain; other DWGs are closed.
        - You can choose which one is active via ``active`` ("a" or "b").

        Parameters
        - ``doc_a`` / ``doc_b``: absolute paths to target DWGs (will be opened if not already opened).
        - ``active``: optional string starting with "a" or "b" to control which doc is activated.

        Returns
        - (basename_a, basename_b) if success (and the two documents are present), otherwise ``()``.

        Usage
        - From Python: ``CadController().standardize_two_documents(a, b, active='a')``
        - From CLI: ``python cad_cli.py standard2 <A.dwg> <B.dwg> [a|b]``
        """
        self.ensure_single_process()
        if not self.start_tarch_v9():
            return tuple()
        self._ensure_li()

        # 2) open targets (idempotent)
        a = self._normalize_path(doc_a)
        b = self._normalize_path(doc_b)
        self.open_dwg(a)
        self.open_dwg(b)
        # Close any error-report windows that might appear after opening
        self._close_error_windows()

        # 3) close others
        try:
            self.wait_quiescent(timeout=30)
            for _ in range(10):
                removed = False
                for d in list(self.ops.acad.Documents):
                    try:
                        full = getattr(d, "FullName", None)
                        nf = self._normalize_path(full) if full else ""
                        if nf and nf not in (a, b):
                            d.Close(False)
                            if full:
                                self.wait_document_closed(full, timeout=60)
                            removed = True
                            break
                    except Exception:
                        pass
                if not removed:
                    break
                self.wait_quiescent(timeout=15)
        finally:
            self.wait_quiescent(timeout=30)

        # 4) activate requested target
        if active:
            want = a if str(active).lower().startswith("a") else b
            self._activate_document(want)

        # 5) UI nicety
        self.activate_cad_window(title_key="AutoCAD", click_titlebar=True, minimize_others=True)
        time.sleep(wait_s)
        self._close_error_windows()

        # 6) collect basenames and return
        try:
            docs = list(self.ops.acad.Documents)
            names = [str(getattr(d, "Name", "")) for d in docs]
        except Exception:
            names = []
        if len(names) == 2:
            print(f"[double] {names[0]}, {names[1]}")
            return (names[0], names[1])
        return tuple()
    def standardize_two_default(self, active: Optional[str] = None) -> tuple[str, str]:
        """Convenience: two-file definite state using built-in default DWGs.

        - Uses ``artifacts/dwgs/00.dwg`` and ``artifacts/dwgs/01.dwg`` under the task folder.
        - Helpful for smoke tests and demos.

        CLI shortcut: ``python cad_cli.py standard2`` (without parameters).
        """

        root = Path(__file__).resolve().parents[1]
        a = str((root / "artifacts" / "dwgs" / "00.dwg").resolve())
        b = str((root / "artifacts" / "dwgs" / "01.dwg").resolve())
        return self.standardize_two_documents(a, b, active=active)




