"""超时保护基建：给同步调用加超时上限，防止 COM/大文件处理无限挂起。

模式与 web-search-server 的 `_with_timeout` 一致：模块级常驻线程池 +
``future.result(timeout)``，超时后丢弃 future（不阻塞等待、不中断后台任务），
由调用方决定返回哪个默认结果。Office/COM 这类可能弹窗或未响应的同步操作
返回一个明确的错误结果，而不是让整个 MCP 调用永久阻塞。
"""

from __future__ import annotations

import concurrent.futures
import threading
from typing import Callable, TypeVar

_T = TypeVar("_T")

_executor: concurrent.futures.ThreadPoolExecutor | None = None
_executor_lock = threading.Lock()


def _get_executor() -> concurrent.futures.ThreadPoolExecutor:
    """懒初始化并返回模块级常驻线程池。"""
    global _executor
    if _executor is None:
        with _executor_lock:
            if _executor is None:
                _executor = concurrent.futures.ThreadPoolExecutor(
                    max_workers=4, thread_name_prefix="office-writer"
                )
    return _executor


def _discard_future(future: concurrent.futures.Future) -> None:
    """丢弃超时 future：取消未启动的任务，并消费已完成任务的异常避免警告。"""
    future.cancel()

    def _consume(f: concurrent.futures.Future) -> None:
        try:
            f.exception()
        except Exception:
            pass

    future.add_done_callback(_consume)


def with_timeout(
    func: Callable[..., _T],
    timeout: float,
    default_result: _T,
    *args,
    **kwargs,
) -> _T:
    """在常驻线程池中执行 ``func(*args, **kwargs)``，超时返回 ``default_result``。

    函数内部抛出的异常照常传播（与无超时时行为一致）。
    """
    future = _get_executor().submit(func, *args, **kwargs)
    try:
        return future.result(timeout=timeout)
    except concurrent.futures.TimeoutError:
        _discard_future(future)
        return default_result
