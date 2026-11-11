"""Apply hard Windows job commit limits before a benchmark process can execute."""
import ctypes as ct
from ctypes import wintypes as w
from pathlib import Path
import sys


def load_owned_job(directory):
    sys.path.insert(0, str(Path(directory)))
    from owned_process import OwnedJob, ExtendedLimits
    return OwnedJob, ExtendedLimits


class CompletionPort(ct.Structure):
    _fields_ = [("key", ct.c_void_p), ("port", w.HANDLE)]


class Accounting(ct.Structure):
    _fields_ = [(name, ct.c_int64) for name in ("user", "kernel", "period_user", "period_kernel")] + [
        (name, w.DWORD) for name in ("faults", "processes", "active", "terminated")]


class MemoryJob:
    def __init__(self, helper_directory, limit_bytes):
        owned_job, self.limits_type = load_owned_job(helper_directory)
        self.job = owned_job()
        self.kernel = self.job.kernel
        self.kernel.OpenProcess.argtypes = [w.DWORD, w.BOOL, w.DWORD]
        self.kernel.OpenProcess.restype = w.HANDLE
        self.kernel.QueryInformationJobObject.argtypes = [w.HANDLE, ct.c_int, ct.c_void_p, w.DWORD, ct.c_void_p]
        self.kernel.CreateIoCompletionPort.argtypes = [w.HANDLE, w.HANDLE, ct.c_size_t, w.DWORD]
        self.kernel.CreateIoCompletionPort.restype = w.HANDLE
        self.kernel.GetQueuedCompletionStatus.argtypes = [w.HANDLE, ct.POINTER(w.DWORD),
                                                         ct.POINTER(ct.c_size_t), ct.POINTER(ct.c_void_p), w.DWORD]
        self.port = self.kernel.CreateIoCompletionPort(ct.c_void_p(-1).value, None, 0, 1)
        if not self.port:
            raise ct.WinError(ct.get_last_error())
        association = CompletionPort(1, self.port)
        self.set_information(7, association)
        limits = self.limits_type()
        limits.basic.flags = 0x2000 | 0x200
        limits.memory[1] = limit_bytes
        self.set_information(9, limits)

    def set_information(self, kind, value):
        if not self.kernel.SetInformationJobObject(self.job.handle, kind, ct.byref(value), ct.sizeof(value)):
            raise ct.WinError(ct.get_last_error())

    def spawn(self, command, **kwargs):
        return self.job.spawn(command, **kwargs)

    def attach(self, pid):
        handle = self.kernel.OpenProcess(0x0100 | 0x0001, False, pid)
        if not handle:
            raise ct.WinError(ct.get_last_error())
        try:
            if not self.kernel.AssignProcessToJobObject(self.job.handle, handle):
                raise ct.WinError(ct.get_last_error())
        finally:
            self.kernel.CloseHandle(handle)

    def memory_events(self):
        events = []
        while True:
            message, key, process = w.DWORD(), ct.c_size_t(), ct.c_void_p()
            if not self.kernel.GetQueuedCompletionStatus(self.port, ct.byref(message), ct.byref(key), ct.byref(process), 0):
                break
            if message.value in (9, 10):
                events.append({"message": message.value, "pid": process.value})
        return events

    def process_ids(self):
        buffer = ct.create_string_buffer(8 + 128 * ct.sizeof(ct.c_size_t))
        if not self.kernel.QueryInformationJobObject(self.job.handle, 3, buffer, ct.sizeof(buffer), None):
            raise ct.WinError(ct.get_last_error())
        count = w.DWORD.from_buffer(buffer, 4).value
        return list((ct.c_size_t * count).from_buffer(buffer, 8))

    def counters(self):
        limits, accounting = self.limits_type(), Accounting()
        for kind, value in ((9, limits), (1, accounting)):
            if not self.kernel.QueryInformationJobObject(self.job.handle, kind, ct.byref(value), ct.sizeof(value), None):
                raise ct.WinError(ct.get_last_error())
        return {"peak_job_committed_bytes": limits.memory[3],
                "peak_process_committed_bytes": limits.memory[2],
                "cpu_user_seconds": accounting.user / 10_000_000,
                "cpu_kernel_seconds": accounting.kernel / 10_000_000,
                "active_processes": accounting.active}

    def terminate(self):
        self.job.terminate()

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.job.close()
        self.kernel.CloseHandle(self.port)
