"""Hold verified process handles when replacing only the corpus supervisor."""
import ctypes
from ctypes import wintypes
import subprocess

k=ctypes.WinDLL('kernel32',use_last_error=True)
k.OpenProcess.argtypes=[wintypes.DWORD,wintypes.BOOL,wintypes.DWORD];k.OpenProcess.restype=wintypes.HANDLE
k.GetProcessTimes.argtypes=[wintypes.HANDLE]+[ctypes.POINTER(wintypes.FILETIME)]*4
k.GetExitCodeProcess.argtypes=[wintypes.HANDLE,ctypes.POINTER(wintypes.DWORD)]
k.WaitForSingleObject.argtypes=[wintypes.HANDLE,wintypes.DWORD];k.WaitForSingleObject.restype=wintypes.DWORD
k.TerminateProcess.argtypes=[wintypes.HANDLE,wintypes.UINT]
k.CloseHandle.argtypes=[wintypes.HANDLE]


class ExistingProcess:
    def __init__(self,pid,created_filetime):
        self.pid=int(pid);self.handle=None;self.returncode=None
        handle=k.OpenProcess(0x100000|0x1000|0x1,False,self.pid)
        if not handle:
            if ctypes.get_last_error()==87:self.returncode=1;return
            raise ctypes.WinError(ctypes.get_last_error())
        self.handle=handle
        times=[wintypes.FILETIME() for _ in range(4)]
        if not k.GetProcessTimes(handle,*[ctypes.byref(t) for t in times]):raise ctypes.WinError(ctypes.get_last_error())
        actual=(times[0].dwHighDateTime<<32)|times[0].dwLowDateTime
        if str(actual)!=str(created_filetime):
            k.CloseHandle(handle);self.handle=None;self.returncode=1

    def poll(self):
        if not self.handle:return self.returncode
        state=k.WaitForSingleObject(self.handle,0)
        if state==0x102:return None
        if state!=0:raise ctypes.WinError(ctypes.get_last_error())
        code=wintypes.DWORD()
        if not k.GetExitCodeProcess(self.handle,ctypes.byref(code)):raise ctypes.WinError(ctypes.get_last_error())
        self.returncode=code.value
        return self.returncode

    def terminate(self):
        if self.poll() is None and not k.TerminateProcess(self.handle,1):raise ctypes.WinError(ctypes.get_last_error())

    def wait(self,timeout=None):
        if self.handle and k.WaitForSingleObject(self.handle,0xffffffff if timeout is None else int(timeout*1000))==0x102:
            raise subprocess.TimeoutExpired(str(self.pid),timeout)
        return self.poll()

    def __del__(self):
        if self.handle:k.CloseHandle(self.handle)
