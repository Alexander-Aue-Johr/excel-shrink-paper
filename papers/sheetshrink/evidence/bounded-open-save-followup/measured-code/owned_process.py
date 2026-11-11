"""Contain a launched command and its descendants in a Windows job."""
import ctypes as ct
from ctypes import wintypes as w
import subprocess


class BasicLimits(ct.Structure):
    _fields_ = [("process_time", ct.c_int64), ("job_time", ct.c_int64),
                ("flags", w.DWORD), ("min_ws", ct.c_size_t), ("max_ws", ct.c_size_t),
                ("active_limit", w.DWORD), ("affinity", ct.c_size_t),
                ("priority", w.DWORD), ("scheduling", w.DWORD)]


class ExtendedLimits(ct.Structure):
    _fields_ = [("basic", BasicLimits), ("io", ct.c_uint64 * 6),
                ("memory", ct.c_size_t * 4)]


def creation_ticks(pid):
    kernel = ct.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = [w.DWORD, w.BOOL, w.DWORD]
    kernel.OpenProcess.restype = w.HANDLE
    kernel.CloseHandle.argtypes = [w.HANDLE]
    kernel.GetProcessTimes.argtypes = [w.HANDLE] + [ct.POINTER(w.FILETIME)] * 4
    handle = kernel.OpenProcess(0x1000, False, pid)
    if not handle:
        error = ct.get_last_error()
        if error == 87:
            return None
        raise ct.WinError(error)
    try:
        ticks = [w.FILETIME() for _ in range(4)]
        if not kernel.GetProcessTimes(handle, *(ct.byref(value) for value in ticks)):
            raise ct.WinError(ct.get_last_error())
        if ticks[1].dwHighDateTime or ticks[1].dwLowDateTime:
            return None
        return (ticks[0].dwHighDateTime << 32) | ticks[0].dwLowDateTime
    finally:
        kernel.CloseHandle(handle)


def terminate_if_same(pid, expected_ticks):
    kernel = ct.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = [w.DWORD, w.BOOL, w.DWORD]
    kernel.OpenProcess.restype = w.HANDLE
    kernel.CloseHandle.argtypes = [w.HANDLE]
    kernel.GetProcessTimes.argtypes = [w.HANDLE] + [ct.POINTER(w.FILETIME)] * 4
    kernel.TerminateProcess.argtypes = [w.HANDLE, w.UINT]
    handle = kernel.OpenProcess(0x1001, False, pid)
    if not handle:
        raise ct.WinError(ct.get_last_error())
    try:
        ticks = [w.FILETIME() for _ in range(4)]
        if not kernel.GetProcessTimes(handle, *(ct.byref(value) for value in ticks)):
            raise ct.WinError(ct.get_last_error())
        actual_ticks = (ticks[0].dwHighDateTime << 32) | ticks[0].dwLowDateTime
        if actual_ticks != expected_ticks:
            raise RuntimeError("Process identity changed before termination")
        if not kernel.TerminateProcess(handle, 0):
            raise ct.WinError(ct.get_last_error())
    finally:
        kernel.CloseHandle(handle)


class OwnedJob:
    def __init__(self):
        self.kernel = ct.WinDLL("kernel32", use_last_error=True)
        self.kernel.CreateJobObjectW.argtypes = [ct.c_void_p, w.LPCWSTR]
        self.kernel.CreateJobObjectW.restype = w.HANDLE
        self.kernel.SetInformationJobObject.argtypes = [w.HANDLE, ct.c_int, ct.c_void_p, w.DWORD]
        self.kernel.AssignProcessToJobObject.argtypes = [w.HANDLE, w.HANDLE]
        self.kernel.TerminateJobObject.argtypes = [w.HANDLE, w.UINT]
        self.kernel.CloseHandle.argtypes = [w.HANDLE]
        self.resume = ct.WinDLL("ntdll").NtResumeProcess
        self.resume.argtypes = [w.HANDLE]
        self.handle = self.kernel.CreateJobObjectW(None, None)
        if not self.handle:
            raise ct.WinError(ct.get_last_error())
        limits = ExtendedLimits()
        limits.basic.flags = 0x2000
        if not self.kernel.SetInformationJobObject(self.handle, 9, ct.byref(limits), ct.sizeof(limits)):
            self.close()
            raise ct.WinError(ct.get_last_error())

    def spawn(self, command, **kwargs):
        process = subprocess.Popen(command, creationflags=0x08000004, **kwargs)
        try:
            if not self.kernel.AssignProcessToJobObject(self.handle, int(process._handle)):
                raise ct.WinError(ct.get_last_error())
            if self.resume(int(process._handle)) != 0:
                raise OSError("Could not resume the contained validator")
            return process
        except BaseException:
            process.kill()
            process.wait()
            raise

    def terminate(self):
        if not self.kernel.TerminateJobObject(self.handle, 124):
            raise ct.WinError(ct.get_last_error())

    def close(self):
        self.kernel.CloseHandle(self.handle)

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()

