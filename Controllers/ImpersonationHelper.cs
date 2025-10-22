using System;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.ComponentModel;

public class ImpersonationHelper : IDisposable
{
    [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
    private static extern bool LogonUser(string username, string domain, string password,
        int logonType, int logonProvider, out IntPtr tokenHandle);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
    private static extern bool CloseHandle(IntPtr handle);

    private WindowsImpersonationContext _impersonationContext;
    private IntPtr _userHandle = IntPtr.Zero;

    private const int LOGON32_LOGON_NEW_CREDENTIALS = 9;
    private const int LOGON32_PROVIDER_DEFAULT = 0;

    public ImpersonationHelper(string domain, string username, string password)
    {
        bool success = LogonUser(username, domain, password,
            LOGON32_LOGON_NEW_CREDENTIALS, LOGON32_PROVIDER_DEFAULT,
            out _userHandle);

        if (!success)
        {
            int err = Marshal.GetLastWin32Error();
            throw new Win32Exception(err, $"LogonUser failed for {domain}\\{username}");
        }

        var newId = new WindowsIdentity(_userHandle);
        _impersonationContext = newId.Impersonate();
    }

    public void Dispose()
    {
        _impersonationContext?.Undo();
        if (_userHandle != IntPtr.Zero)
        {
            CloseHandle(_userHandle);
            _userHandle = IntPtr.Zero;
        }
    }
}
