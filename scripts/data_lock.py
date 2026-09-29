# -*- coding: utf-8 -*-
"""
全局数据文件锁：串行化所有写端点，防多用户并发写坏 Excel/JSON（方案A-P0）。

实现：独占锁文件 + OS 级字节锁（Windows msvcrt / POSIX fcntl），带超时。
用法：
    with data_lock(timeout=180):
        ... 读-改-写 ...
"""
import os
import time

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOCK_FILE = os.path.join(BASE, "output", ".data.lock")

_is_win = os.name == "nt"
if _is_win:
    import msvcrt

    def _lock_byte(f):
        # NBLCK 立即失败，由外层循环控制重试节奏与超时
        # （LK_LOCK 自带 10s 内部重试会掩盖自定义 timeout）
        msvcrt.locking(f.fileno(), msvcrt.LK_NBLCK, 1)

    def _unlock_byte(f):
        f.seek(0)
        msvcrt.locking(f.fileno(), msvcrt.LK_UNLCK, 1)
else:
    import fcntl

    def _lock_byte(f):
        fcntl.flock(f.fileno(), fcntl.LOCK_EX)

    def _unlock_byte(f):
        fcntl.flock(f.fileno(), fcntl.LOCK_UN)


class _DataLock:
    """with 语义；timeout 秒拿不到锁抛 TimeoutError。同进程同线程可重入。"""
    _held = None  # (thread_id, fh) 重入检测（flask before_request 与视图层可能嵌套）

    def __init__(self, path=LOCK_FILE, timeout=180):
        self.path = path
        self.timeout = timeout
        import threading
        self._tid = threading.get_ident()

    def __enter__(self):
        if _DataLock._held is not None and _DataLock._held[0] == self._tid:
            self._reentered = True
            return self
        self._reentered = False
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        f = open(self.path, "a+b")
        start = time.time()
        while True:
            try:
                f.seek(0)
                _lock_byte(f)
                break
            except OSError:
                if time.time() - start > self.timeout:
                    f.close()
                    raise TimeoutError("数据被另一操作锁定超过 %d 秒（导入/对齐可能正在运行），请稍后重试" % self.timeout)
                time.sleep(0.4)
        _DataLock._held = (self._tid, f)
        self._f = f
        return self

    def __exit__(self, exc_type, exc, tb):
        if self._reentered:
            return False
        try:
            _unlock_byte(self._f)
        finally:
            self._f.close()
            _DataLock._held = None
        return False


def data_lock(timeout=180):
    return _DataLock(timeout=timeout)
