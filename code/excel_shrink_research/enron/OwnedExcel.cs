using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

// Excel is assigned to a private job before its first instruction executes.
// A blocked COM call cannot prevent cleanup when the supervisor ends this worker.
public sealed class OwnedExcel : IDisposable
{
    private IntPtr job;
    private IntPtr process;
    public int Pid { get; private set; }
    public string CreatedFileTime { get; private set; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct StartupInfo
    {
        public int cb; public string reserved; public string desktop; public string title;
        public int x; public int y; public int xSize; public int ySize;
        public int xCountChars; public int yCountChars; public int fillAttribute;
        public int flags; public short showWindow; public short reserved2;
        public IntPtr reservedPointer; public IntPtr stdInput; public IntPtr stdOutput; public IntPtr stdError;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct ProcessInformation { public IntPtr process; public IntPtr thread; public int pid; public int tid; }
    [StructLayout(LayoutKind.Sequential)]
    private struct BasicLimitInformation
    {
        public long processTime; public long jobTime; public uint flags;
        public UIntPtr minWorkingSet; public UIntPtr maxWorkingSet; public uint activeProcesses;
        public UIntPtr affinity; public uint priorityClass; public uint schedulingClass;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct IoCounters { public ulong a; public ulong b; public ulong c; public ulong d; public ulong e; public ulong f; }
    [StructLayout(LayoutKind.Sequential)]
    private struct ExtendedLimitInformation
    {
        public BasicLimitInformation basic; public IoCounters io;
        public UIntPtr processMemory; public UIntPtr jobMemory; public UIntPtr peakProcessMemory; public UIntPtr peakJobMemory;
    }
    private delegate bool WindowCallback(IntPtr window, IntPtr parameter);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateJobObject(IntPtr attributes, string name);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool SetInformationJobObject(IntPtr job, int infoClass, ref ExtendedLimitInformation info, uint length);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AssignProcessToJobObject(IntPtr job, IntPtr process);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool CreateProcess(string application, StringBuilder commandLine, IntPtr processAttributes,
        IntPtr threadAttributes, bool inherit, uint flags, IntPtr environment, string directory,
        ref StartupInfo startup, out ProcessInformation info);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern uint ResumeThread(IntPtr thread);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll")] private static extern bool TerminateProcess(IntPtr process, uint code);
    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GetProcessTimes(IntPtr process, out long creation, out long exit, out long kernel, out long user);
    [DllImport("user32.dll")] private static extern bool EnumWindows(WindowCallback callback, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumChildWindows(IntPtr parent, WindowCallback callback, IntPtr parameter);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int size);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("oleacc.dll")] private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid iid,
        [MarshalAs(UnmanagedType.Interface)] out object result);

    public OwnedExcel(string executable, string seed)
    {
        job = CreateJobObject(IntPtr.Zero, null);
        if (job == IntPtr.Zero) throw new Win32Exception();
        ExtendedLimitInformation limits = new ExtendedLimitInformation();
        limits.basic.flags = 0x2000; // JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        if (!SetInformationJobObject(job, 9, ref limits, (uint)Marshal.SizeOf(limits)))
        { Dispose(); throw new Win32Exception(); }
        StartupInfo startup = new StartupInfo();
        startup.cb = Marshal.SizeOf(startup);
        startup.flags = 1; // STARTF_USESHOWWINDOW
        startup.showWindow = 0;
        ProcessInformation info;
        // Both paths originate from local filesystem paths, which cannot contain quotes.
        StringBuilder command = new StringBuilder("\"" + executable + "\" /x /e \"" + seed + "\"");
        if (!CreateProcess(executable, command, IntPtr.Zero, IntPtr.Zero, false, 4,
                           IntPtr.Zero, System.IO.Path.GetDirectoryName(seed), ref startup, out info))
        { Dispose(); throw new Win32Exception(); }
        process = info.process;
        Pid = info.pid;
        try
        {
            if (!AssignProcessToJobObject(job, process)) throw new Win32Exception();
            long creation, exit, kernel, user;
            if (!GetProcessTimes(process, out creation, out exit, out kernel, out user)) throw new Win32Exception();
            CreatedFileTime = creation.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (ResumeThread(info.thread) == uint.MaxValue) throw new Win32Exception();
        }
        catch { TerminateProcess(process, 125); Dispose(); throw; }
        finally { CloseHandle(info.thread); }
    }

    public object FindNativeWindow()
    {
        object native = null;
        EnumWindows(delegate(IntPtr top, IntPtr unused)
        {
            uint pid;
            GetWindowThreadProcessId(top, out pid);
            if (pid != (uint)Pid) return true;
            EnumChildWindows(top, delegate(IntPtr child, IntPtr unusedChild)
            {
                StringBuilder name = new StringBuilder(128);
                GetClassName(child, name, name.Capacity);
                if (name.ToString() != "EXCEL7") return true;
                Guid dispatch = new Guid("00020400-0000-0000-C000-000000000046");
                object candidate;
                if (AccessibleObjectFromWindow(child, 0xFFFFFFF0, ref dispatch, out candidate) == 0)
                { native = candidate; return false; }
                return true;
            }, IntPtr.Zero);
            return native == null;
        }, IntPtr.Zero);
        return native;
    }

    public void Dispose()
    {
        if (job != IntPtr.Zero) { CloseHandle(job); job = IntPtr.Zero; }
        if (process != IntPtr.Zero) { CloseHandle(process); process = IntPtr.Zero; }
    }
}
