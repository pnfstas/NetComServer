using Microsoft.Win32;
using System.Diagnostics;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading;
using System.Xml;

using static NetComServer.Helpers.Win32Helper;
using static NetComServer.Helpers.RegistryHelper;

namespace NetComServer.Helpers
{
    public delegate void WriteLineDelegate(string? message);
    public class Win32Helper
    {
        #region WinErrors
		public const int S_OK							= 0;
		public const int S_FALSE						= 1;
		public const int ERROR_ACCESS_DENIED			= 5;
		public const int ERROR_INVALID_HANDLE			= 6;
		public const int ERROR_GEN_FAILURE				= 31;
		public const int ERROR_MORE_DATA				= 234;
		public const int ERROR_NO_SUCH_ALIAS			= 1376;
		public const int E_NOTIMPL						= unchecked((int)2147500033);
		public const int E_NOINTERFACE					= unchecked((int)2147500034);
		public const int E_POINTER						= unchecked((int)2147500035);
		public const int E_ABORT						= unchecked((int)2147500036);
		public const int E_FAIL							= unchecked((int)2147500037);
		public const int E_UNEXPECTED					= unchecked((int)2147549183);
		public const int CLASS_E_NOAGGREGATION			= unchecked((int)2147746064);
		public const int CLASS_E_CLASSNOTAVAILABLE		= unchecked((int)2147746065);
		public const int REGDB_E_READREGDB				= unchecked((int)2147746128);
		public const int REGDB_E_CLASSNOTREG			= unchecked((int)2147746132);
		public const int CO_E_APPNOTFOUND				= unchecked((int)2147746293);
		public const int CO_E_DLLNOTFOUND				= unchecked((int)2147746296);
		public const int CO_E_ERRORINDLL				= unchecked((int)2147746297);
		public const int CO_E_OBJISREG					= unchecked((int)2147746300);
		public const int CO_E_APPDIDNTREG				= unchecked((int)2147746302);
		public const int E_ACCESSDENIED					= unchecked((int)2147942405);
		public const int E_HANDLE						= unchecked((int)2147942406);
		public const int E_OUTOFMEMORY					= unchecked((int)2147942414);
		public const int E_INVALIDARG					= unchecked((int)2147942487);
		public const int CO_E_CLASS_CREATE_FAILED		= unchecked((int)2148007937);
		public const int CO_E_SCM_ERROR					= unchecked((int)2148007938);
		public const int CO_E_SCM_RPC_FAILURE			= unchecked((int)2148007939);
		public const int CO_E_BAD_PATH					= unchecked((int)2148007940);
		public const int CO_E_SERVER_EXEC_FAILURE		= unchecked((int)2148007941);
		public const int CO_E_OBJSRV_RPC_FAILURE		= unchecked((int)2148007942);
		public const int MK_E_NO_NORMALIZED				= unchecked((int)2148007943);
		public const int CO_E_SERVER_STOPPING			= unchecked((int)2148007944);
		public const int NERR_Success					= 0;
		public const int NERR_BASE						= 2100;
		public const int NERR_GroupNotFound				= NERR_BASE + 120;		//The group name could not be found.
		public const int NERR_UserNotFound				= NERR_BASE + 121;		//The user name could not be found.
		public const int NERR_ResourceNotFound			= NERR_BASE + 122;		//The resource name could not be found.
		public const int NERR_GroupExists				= NERR_BASE + 123;		//The group already exists.
		public const int NERR_UserExists				= NERR_BASE+124;		//The account already exists.
		public const int NERR_InvalidComputer			= NERR_BASE + 251;
		public const uint MAX_PREFERRED_LENGTH			= unchecked((uint)-1);
        #endregion
        #region Guids
        public const string IID_IClassFactory = "00000001-0000-0000-C000-000000000046";
        public const string IID_IUnknown = "00000000-0000-0000-C000-000000000046";
        public const string IID_IDispatch = "00020400-0000-0000-C000-000000000046";
        #endregion
        public enum CLSCTX
        {
            INPROC_SERVER = 0x1,
            INPROC_HANDLER = 0x2,
            LOCAL_SERVER = 0x4,
            INPROC_SERVER16 = 0x8,
            REMOTE_SERVER = 0x10,
            INPROC_HANDLER16 = 0x20,
            RESERVED1 = 0x40,
            RESERVED2 = 0x80,
            RESERVED3 = 0x100,
            RESERVED4 = 0x200,
            NO_CODE_DOWNLOAD = 0x400,
            RESERVED5 = 0x800,
            NO_CUSTOM_MARSHAL = 0x1000,
            ENABLE_CODE_DOWNLOAD = 0x2000,
            NO_FAILURE_LOG = 0x4000,
            DISABLE_AAA = 0x8000,
            ENABLE_AAA = 0x10000,
            FROM_DEFAULT_CONTEXT = 0x20000,
            ACTIVATE_X86_SERVER = 0x40000,
            ACTIVATE_32_BIT_SERVER = ACTIVATE_X86_SERVER,
            ACTIVATE_64_BIT_SERVER = 0x80000,
            ENABLE_CLOAKING = 0x100000,
            APPCONTAINER = 0x400000,
            ACTIVATE_AAA_AS_IU = 0x800000,
            RESERVED6 = 0x1000000,
            ACTIVATE_ARM32_SERVER = 0x2000000,
            ALLOW_LOWER_TRUST_REGISTRATION = 0x4000000,
            PS_DLL = unchecked((int)0x80000000)
        }
        public enum REGCLS
        {
            SINGLEUSE = 0,
            MULTIPLEUSE = 1,
            MULTI_SEPARATE = 2,
            SUSPENDED = 4,
            SURROGATE = 8,
            AGILE = 0x10,
        }
        public enum CMDSHOW
        {
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_FORCEMINIMIZE = 11,
            SW_MAX = 11
        }
        public enum MessageBoxFlag : uint
        {
            MB_OK                       = 0x00000000,
            MB_OKCANCEL                 = 0x00000001,
            MB_ABORTRETRYIGNORE         = 0x00000002,
            MB_YESNOCANCEL              = 0x00000003,
            MB_YESNO                    = 0x00000004,
            MB_RETRYCANCEL              = 0x00000005,
            MB_CANCELTRYCONTINUE        = 0x00000006,
            MB_ICONHAND                 = 0x00000010,
            MB_ICONQUESTION             = 0x00000020,
            MB_ICONEXCLAMATION          = 0x00000030,
            MB_ICONASTERISK             = 0x00000040,
            MB_USERICON                 = 0x00000080,
            MB_ICONWARNING              = MB_ICONEXCLAMATION,
            MB_ICONERROR                = MB_ICONHAND,
            MB_ICONINFORMATION          = MB_ICONASTERISK,
            MB_ICONSTOP                 = MB_ICONHAND,
            MB_DEFBUTTON1               = 0x00000000,
            MB_DEFBUTTON2               = 0x00000100,
            MB_DEFBUTTON3               = 0x00000200,
            MB_DEFBUTTON4               = 0x00000300,
            MB_APPLMODAL                = 0x00000000,
            MB_SYSTEMMODAL              = 0x00001000,
            MB_TASKMODAL                = 0x00002000,
            MB_HELP                     = 0x00004000,
            MB_NOFOCUS                  = 0x00008000,
            MB_SETFOREGROUND            = 0x00010000,
            MB_DEFAULT_DESKTOP_ONLY     = 0x00020000,
            MB_TOPMOST                  = 0x00040000,
            MB_RIGHT                    = 0x00080000,
            MB_RTLREADING               = 0x00100000,
            MB_SERVICE_NOTIFICATION     = 0x00200000,
            MB_TYPEMASK                 = 0x0000000F,
            MB_ICONMASK                 = 0x000000F0,
            MB_DEFMASK                  = 0x00000F00,
            MB_MODEMASK                 = 0x00003000,
            MB_MISCMASK                 = 0x0000C000
        }
        public enum DialogBoxCommandId
        {
            IDOK            = 1,
            IDCANCEL        = 2,
            IDABORT         = 3,
            IDRETRY         = 4,
            IDIGNORE        = 5,
            IDYES           = 6,
            IDNO            = 7,
            IDCLOSE         = 8,
            IDHELP          = 9,
            IDTRYAGAIN      = 10,
            IDCONTINUE      = 11,
            IDTIMEOUT       = 32000
        }
        public enum REGKIND
        {
            DEFAULT = 0,
            REGISTER,
            NONE
        }
        public enum COM_RIGHTS
        {
            NONE = 0,
            EXECUTE = 1,
            EXECUTE_LOCAL = 2,
            EXECUTE_REMOTE = 4,
            ACTIVATE_LOCAL = 8,
            ACTIVATE_REMOTE = 16
        }
        public enum SID_NAME_USE
        {
            SidTypeUser = 1,
            SidTypeGroup,
            SidTypeDomain,
            SidTypeAlias,
            SidTypeWellKnownGroup,
            SidTypeDeletedAccount,
            SidTypeInvalid,
            SidTypeUnknown,
            SidTypeComputer,
            SidTypeLabel,
            SidTypeLogonSession
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public nint hwnd;
            public uint message;
            public nint wParam;
            public nint lParam;
            public uint time;
            public POINT pt;
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }
		[StructLayout(LayoutKind.Sequential)]
		public struct LOCALGROUP_MEMBERS_INFO_2
		{
			public nint Sid;
			public SID_NAME_USE SidUsage;
			[MarshalAs(UnmanagedType.LPWStr)] public string DomainAndName;
		}
        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid(IID_IClassFactory)]
        public interface IClassFactory
        {
            [PreserveSig]
            abstract int CreateInstance([In, MarshalAs(UnmanagedType.IUnknown)] IntPtr pUnkOuter, [In] ref Guid riid, [Out] out IntPtr ppvObject);
            [PreserveSig]
            abstract int LockServer([In] bool fLock);
        }
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetMessage(
            [Out] out MSG lpMsg,
            [In, Optional] nint hWnd,
            [In] uint wMsgFilterMin,
            [In] uint wMsgFilterMax);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool TranslateMessage([In] ref MSG lpMsg);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern long DispatchMessage([In] ref MSG lpMsg);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern nint GetForegroundWindow();
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int GetWindowThreadProcessId(
            [In] nint hWnd,
            [Out] out int dwProcessId);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool ShowWindow(
            [In] nint hWnd,
            [In] CMDSHOW nCmdShow);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int MessageBox(
            [In, Optional] nint hWnd,
            [In, Optional, MarshalAs(UnmanagedType.LPTStr)] string lpText,
            [In, Optional, MarshalAs(UnmanagedType.LPTStr)] string lpCaption,
            [In, MarshalAs(UnmanagedType.U4)] MessageBoxFlag uType);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern nint GetConsoleWindow();
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool AllocConsole();
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool AttachConsole([In] int dwProcessId);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool FreeConsole();
        [DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int CoRegisterClassObject(
            [In, MarshalAs(UnmanagedType.LPStruct)] ref Guid rclsid,
            [In, MarshalAs(UnmanagedType.IUnknown)] object pUnk,
            [In, MarshalAs(UnmanagedType.U4)] CLSCTX dwClsContext,
            [In, MarshalAs(UnmanagedType.U4)] REGCLS flags,
            [Out, MarshalAs(UnmanagedType.U4)] out uint lpdwRegister);
        [DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int CoResumeClassObjects();
        [DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int CoRevokeClassObject([In] uint dwRegister);
        [DllImport("ole32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int CoSuspendClassObjects();
        [DllImport("oleaut32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int LoadTypeLibEx(
            [In, MarshalAs(UnmanagedType.LPTStr)] string szFile,
            [In] REGKIND regkind,
            [Out] out ITypeLib pptlib);
        [DllImport("oleaut32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int LoadRegTypeLib(
            [In, Out] ref Guid rguid,
            [In] ushort wVerMajor,
            [In] ushort wVerMinor,
            [In] int lcid,
            [Out] out ITypeLib pptlib);
        [DllImport("oleaut32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int UnRegisterTypeLib(
            [In, Out] ref Guid rguid,
            [In] ushort wVerMajor,
            [In] ushort wVerMinor,
            [In] int lcid,
            [In] SYSKIND syskind);
		[DllImport("netapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
		public static extern int NetLocalGroupGetMembers(
			[In, MarshalAs(UnmanagedType.LPWStr)] string servername,
			[In, MarshalAs(UnmanagedType.LPWStr)] string localgroupname,
			[In, MarshalAs(UnmanagedType.U4)] uint level,
			[Out] out nint bufptr,
			[In, MarshalAs(UnmanagedType.U4)] uint prefmaxlen,
			[Out, MarshalAs(UnmanagedType.U4)] out uint entriesread,
			[Out, MarshalAs(UnmanagedType.U4)] out uint totalentries,
			[In, Out] ref nint resumehandle
		);
        [DllImport("netapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern int NetApiBufferFree([In] nint Buffer);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public extern static bool CloseHandle([In] IntPtr handle);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public extern static IntPtr LocalFree([In] IntPtr hMem);
        public static bool ShowConsole(bool showMinimized = false)
        {
            bool isShown = false;
            nint hwndConsole = GetConsoleWindow();
            if(hwndConsole != nint.Zero)
            {
                isShown = ShowWindow(hwndConsole, showMinimized ? CMDSHOW.SW_SHOWMINIMIZED : CMDSHOW.SW_SHOW);
            }
            if(!isShown && (hwndConsole == nint.Zero || FreeConsole()))
            {
                nint hwnd = GetForegroundWindow();
                int processId = 0;
                GetWindowThreadProcessId(hwnd, out processId);
                Process process = Process.GetProcessById(processId);
                if(process.ProcessName?.Equals("cmd") == true)
                {
                    isShown = AttachConsole(processId);
                }
                if(!isShown)
                {
                    isShown = AllocConsole();
                }
                if(isShown && showMinimized)
                {
                    hwndConsole = GetConsoleWindow();
                    isShown = hwndConsole != nint.Zero && ShowWindow(hwndConsole, CMDSHOW.SW_MINIMIZE);
                }
            }
            return isShown;
        }
        public static bool IsLocalGroupMember(IdentityReference? group, IdentityReference? identity)
        {
            bool isMember = false;
            NTAccount? groupAccount = group?.Translate(typeof(NTAccount)) as NTAccount;
            SecurityIdentifier? checkedSid = identity?.Translate(typeof(SecurityIdentifier)) as SecurityIdentifier;
            if(groupAccount != null && checkedSid != null)
            {
                uint level = 2;
                nint ptrMembersInfo = nint.Zero;
                uint entriesRead = 0;
                uint totalEntries = 0;
                nint ptrResumeHandle = nint.Zero;
                int result = NetLocalGroupGetMembers(null, groupAccount.Value.Trim(), level, out ptrMembersInfo, MAX_PREFERRED_LENGTH, out entriesRead, out totalEntries, ref ptrResumeHandle);
                if(result == NERR_Success)
                {
                    LOCALGROUP_MEMBERS_INFO_2 membersInfo;
                    nint ptrCurRecord = ptrMembersInfo;
                    int size = Marshal.SizeOf<LOCALGROUP_MEMBERS_INFO_2>();
                    SecurityIdentifier? memberSid = null;
                    string[]? arrGroupNames = null;
                    for(int index = 0; !isMember && index < entriesRead; index++)
                    {
                        ptrCurRecord += index * size;
                        membersInfo = Marshal.PtrToStructure<LOCALGROUP_MEMBERS_INFO_2>(ptrCurRecord);
                        memberSid = new SecurityIdentifier(membersInfo.Sid);
                        if(memberSid == checkedSid)
                        {
                            isMember = true;
                        }
                        else if(membersInfo.SidUsage == SID_NAME_USE.SidTypeGroup || membersInfo.SidUsage == SID_NAME_USE.SidTypeWellKnownGroup
                            || membersInfo.SidUsage == SID_NAME_USE.SidTypeAlias)
                        {
                            arrGroupNames = membersInfo.DomainAndName?.Trim()?.Split('\\', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
                            if(arrGroupNames?.Length > 1)
                            {
                                isMember = IsLocalGroupMember(memberSid, checkedSid);
                            }
                        }
                    }
                }
            }
            return isMember;
        }
    }
	public class RegistryHelper
	{
		public enum RegistryKeyType : int
		{
			Common,
			Compliant
		}
		public class RegistryKeyPair : IEquatable<RegistryKeyPair>, IDisposable
		{
			private bool disposed = false;
			public RegistryKey? Common { get; set; }
			public RegistryKey? Compliant { get; set; }
			public RegistryKeyPair(RegistryKey? common = null, RegistryKey? compliant = null)
			{
				Common = common;
				Compliant = compliant;
			}
			public int KeyCount
			{
				get => Common != null && Compliant != null ? 2 : Common != null || Compliant != null ? 1 : 0;
			}
			public string Name
			{
				get => !string.IsNullOrWhiteSpace(Common?.Name) && !string.IsNullOrWhiteSpace(Compliant?.Name) ?
					$"{{{Common.Name},\n{Compliant?.Name}}}" : !string.IsNullOrWhiteSpace(Common?.Name) ? Common.Name : Compliant?.Name;
			}
			public int SubKeyCount
			{
				get => Math.Max(Common?.SubKeyCount ?? 0, Compliant?.SubKeyCount ?? 0);
			}
			public int ValueCount
			{
				get => Math.Max(Common?.ValueCount ?? 0, Compliant?.ValueCount ?? 0);
			}
			public string[] GetSubKeyNames(RegistryKeyType regKeyType = RegistryKeyType.Common)
			{
				return regKeyType == RegistryKeyType.Compliant ? Compliant?.GetSubKeyNames() : Common?.GetSubKeyNames();
			}
			public RegistryKeyPair? OpenSubKey(string name) => OpenSubKey(name, false);
			public RegistryKeyPair? OpenSubKey(string name, bool writable)
			{
				RegistryKeyPair regKeyPair = new RegistryKeyPair();
				regKeyPair.Common = Common?.OpenSubKey(name, writable);
				regKeyPair.Compliant = Compliant?.OpenSubKey(name, writable);
				return regKeyPair != null ? regKeyPair : null;
			}
			public RegistryKeyPair CreateSubKey(string subkey) => CreateSubKey(subkey, false);
			public RegistryKeyPair CreateSubKey(string subkey, bool readOnly) => CreateSubKey(subkey, readOnly, RegistryOptions.None);
			public RegistryKeyPair CreateSubKey(string subkey, bool readOnly, RegistryOptions options)
			{
				RegistryKeyPair regKeyPair = new RegistryKeyPair();
				regKeyPair.Common = Common?.CreateSubKey(subkey, !readOnly, options);
				regKeyPair.Compliant = Compliant?.CreateSubKey(subkey, !readOnly, options);
				return regKeyPair;
			}
			public void DeleteSubKey(string subkey) => DeleteSubKey(subkey, false);
			public void DeleteSubKey(string subkey, bool throwOnMissingSubKey)
			{
				Common.DeleteSubKey(subkey, throwOnMissingSubKey);
				Compliant.DeleteSubKey(subkey, throwOnMissingSubKey);
			}
			public void DeleteSubKeyTree(string subkey) => DeleteSubKeyTree(subkey, false);
			public void DeleteSubKeyTree(string subkey, bool throwOnMissingSubKey)
			{
				Common.DeleteSubKeyTree(subkey, throwOnMissingSubKey);
				Compliant.DeleteSubKeyTree(subkey, throwOnMissingSubKey);
			}
			public string[] GetValueNames(RegistryKeyType regKeyType = RegistryKeyType.Common)
			{
				return regKeyType == RegistryKeyType.Compliant ? Compliant?.GetValueNames() : Common?.GetValueNames();
			}
			public RegistryValueKind GetValueKind(string? name, RegistryKeyType regKeyType = RegistryKeyType.Common)
			{
				return (regKeyType == RegistryKeyType.Compliant ? Compliant?.GetValueKind(name) : Common?.GetValueKind(name)) ?? RegistryValueKind.Unknown;
			}
			public object? GetValue(string? name, RegistryKeyType regKeyType = RegistryKeyType.Common) => GetValue(name, null, regKeyType);
			public object? GetValue(string? name, object? defaultValue, RegistryKeyType regKeyType = RegistryKeyType.Common) =>
				GetValue(name, defaultValue, RegistryValueOptions.None, regKeyType);
			public object? GetValue(string? name, object? defaultValue, RegistryValueOptions options, RegistryKeyType regKeyType = RegistryKeyType.Common)
			{
				return regKeyType == RegistryKeyType.Compliant ? Compliant?.GetValue(name, defaultValue, options) : Common?.GetValue(name, defaultValue, options);
			}
			public void SetValue(string? name, object value) => SetValue(name, value, RegistryValueKind.Unknown);
			public void SetValue(string? name, object value, RegistryValueKind valueKind)
			{
				Common?.SetValue(name, value, valueKind);
				Compliant?.SetValue(name, value, valueKind);
			}
			public void DeleteValue(string name) => DeleteValue(name, false);
			public void DeleteValue(string name, bool throwOnMissingValue)
			{
				Common.DeleteValue(name, throwOnMissingValue);
				Compliant.DeleteValue(name, throwOnMissingValue);
			}
			public void Flush()
			{
				Common.Flush();
				Compliant.Flush();
			}
			public void Close()
			{
				Common.Close();
				Compliant.Close();
			}
			public void Dispose()
			{
				Dispose(true);
				GC.SuppressFinalize(this);
			}
			protected virtual void Dispose(bool disposing)
			{
				if(!disposed)
				{
					if(disposing)
					{
						Common.Dispose();
						Compliant.Dispose();
					}
					disposed = true;
				}
			}
			public override int GetHashCode()
			{
				return !disposed ? (Common?.GetHashCode() ?? 0) ^ (Compliant?.GetHashCode() ?? 0) : 0;
			}
			public bool Equals(RegistryKeyPair? other)
			{
				return Equals(this, other);
			}
			public override bool Equals(object? other)
			{
				return Equals(this, other);
			}
			public static bool Equals(RegistryKeyPair? regKeyPair1, RegistryKeyPair? regKeyPair2)
			{
				return Object.ReferenceEquals(regKeyPair1, regKeyPair2)
					|| regKeyPair1 is null && (regKeyPair2 is null || regKeyPair2.Common is null || regKeyPair2.Compliant is null)
					|| regKeyPair2 is null && (regKeyPair1 is null || regKeyPair1.Common is null || regKeyPair1.Compliant is null)
					|| regKeyPair1?.Common?.Equals(regKeyPair2?.Common) == true && regKeyPair1?.Compliant?.Equals(regKeyPair2?.Compliant) == true;
			}
			public static bool operator ==(RegistryKeyPair? regKeyPair1, RegistryKeyPair? regKeyPair2)
			{
				return Equals(regKeyPair1, regKeyPair2);
			}
			public static bool operator !=(RegistryKeyPair? regKeyPair1, RegistryKeyPair? regKeyPair2)
			{
				return !Equals(regKeyPair1, regKeyPair2);
			}
			public override string ToString()
			{
				return !disposed && KeyCount > 0 ? "Registry " + (KeyCount > 1 ? "keys " : "key ") + Name : null;
			}
		}
	}
    public class SecurityHelper
    {
        public static RawSecurityDescriptor? GetRawSecurityDescriptor(COM_RIGHTS accessRights, string accountName, byte[] rawCurrentValue = null)
        {
            return GetRawSecurityDescriptor(accessRights, new NTAccount(accountName), rawCurrentValue);
        }
        public static RawSecurityDescriptor? GetRawSecurityDescriptor(COM_RIGHTS rights, IdentityReference identity, byte[] rawCurrentValue = null)
        {
            RawSecurityDescriptor? sd = null;
            try
            {
                SecurityIdentifier? sid = identity?.Translate(typeof(SecurityIdentifier)) as SecurityIdentifier;
                if(sid != null && (int)rights > 0)
                {
                    bool hasAccess = false;
                    RawSecurityDescriptor sdCurrent = null;
                    if(rawCurrentValue?.Length > 0 && (sdCurrent = new RawSecurityDescriptor(rawCurrentValue, 0))?.DiscretionaryAcl?.Count > 0)
                    {
                        hasAccess = //!ExplicitDeny && (ExplicitAllow || !InheritedDeny && InheritedAllow)
                            (from curAce in sdCurrent.DiscretionaryAcl.OfType<CommonAce>()
                             where (curAce.SecurityIdentifier == sid || IsLocalGroupMember(curAce.SecurityIdentifier, sid))
                                && (curAce.AccessMask & (int)rights) >= (int)rights
                                && curAce.AceType == AceType.AccessDenied
                                && !curAce.IsInherited
                             select curAce)?.FirstOrDefault() == null
                            &&
                            ((from curAce in sdCurrent.DiscretionaryAcl.OfType<CommonAce>()
                              where (curAce.SecurityIdentifier == sid || IsLocalGroupMember(curAce.SecurityIdentifier, sid))
                                 && (curAce.AccessMask & (int)rights) >= (int)rights
                                 && curAce.AceType == AceType.AccessAllowed
                                && !curAce.IsInherited
                              select curAce)?.Count() > 0
                            ||
                            (from curAce in sdCurrent.DiscretionaryAcl.OfType<CommonAce>()
                             where (curAce.SecurityIdentifier == sid || IsLocalGroupMember(curAce.SecurityIdentifier, sid))
                                && (curAce.AccessMask & (int)rights) >= (int)rights
                                && curAce.AceType == AceType.AccessDenied
                               && curAce.IsInherited
                             select curAce)?.FirstOrDefault() == null
                            &&
                            (from curAce in sdCurrent.DiscretionaryAcl.OfType<CommonAce>()
                             where (curAce.SecurityIdentifier == sid || IsLocalGroupMember(curAce.SecurityIdentifier, sid))
                                && (curAce.AccessMask & (int)rights) >= (int)rights
                                && curAce.AceType == AceType.AccessAllowed
                               && curAce.IsInherited
                             select curAce)?.Count() > 0);
                    }
                    if(!hasAccess)
                    {
                        if(sdCurrent?.DiscretionaryAcl?.Count > 0)
                        {
                            sd = sdCurrent;
                        }
                        else
                        {
                            SecurityIdentifier sidLocalSystem = new SecurityIdentifier(WellKnownSidType.LocalSystemSid, null);
                            RawAcl acl = new RawAcl(GenericAcl.AclRevision, 1);
                            sd = new RawSecurityDescriptor(ControlFlags.DiscretionaryAclPresent, sidLocalSystem, sidLocalSystem, null, acl);
                        }
                        CommonAce ace = new CommonAce(AceFlags.None, AceQualifier.AccessAllowed, (int)rights, sid, false, null);
                        sd.DiscretionaryAcl?.InsertAce(sd.DiscretionaryAcl.Count, ace);
                    }
                }
            }
            catch(Exception e)
            {
                sd = null;
                Debug.WriteLine(e);
                throw;
            }
            return sd;
        }
        public static bool SetComServerPermissions(RegistryKeyPair regKeyPairAppId, COM_RIGHTS accessRights, COM_RIGHTS launchRights,
            string accessForAccountName = null, string launchAsAccountName = null, WriteLineDelegate writeLineDlgt = null)
        {
            bool isAccessPermissionSet = (int)accessRights == 0;
            bool isLaunchPermissionSet = (int)launchRights == 0;
            bool IsPermissionsSet() => isAccessPermissionSet && isLaunchPermissionSet;
            try
            {
                if(writeLineDlgt == null)
                {
                    writeLineDlgt = new WriteLineDelegate((string? message) => Debug.WriteLine(message));
                }
                if(regKeyPairAppId != null && !IsPermissionsSet())
                {
                    NTAccount accessForAccount = new NTAccount(accessForAccountName);
                    NTAccount launchAsAccount = new NTAccount(launchAsAccountName);
                    SecurityIdentifier accessForSid = accessForAccount?.Translate(typeof(SecurityIdentifier)) as SecurityIdentifier ??
                        WindowsIdentity.GetCurrent().User;
                    SecurityIdentifier launchAsSid = launchAsAccount?.Translate(typeof(SecurityIdentifier)) as SecurityIdentifier ??
                        WindowsIdentity.GetCurrent().User;
                    string FormatPermissions(byte[] permis) => permis?.Length > 0 ? string.Join(' ', permis?.Select((current) => current.ToString("X2"))) : "";
                    if(!isAccessPermissionSet)
                    {
                        byte[] rawAccessPermissions = regKeyPairAppId.GetValue("AccessPermission") as byte[];
                        RawSecurityDescriptor? sdAccessPermissions = GetRawSecurityDescriptor(accessRights, accessForSid, rawAccessPermissions);
                        writeLineDlgt($"{regKeyPairAppId} current value AccessPermission = {FormatPermissions(rawAccessPermissions)}");
                        if(!(isAccessPermissionSet = sdAccessPermissions == null))
                        {
                            rawAccessPermissions = new byte[sdAccessPermissions.BinaryLength];
                            sdAccessPermissions.GetBinaryForm(rawAccessPermissions, 0);
                            regKeyPairAppId.SetValue("AccessPermission", rawAccessPermissions);
                            rawAccessPermissions = regKeyPairAppId.GetValue("AccessPermission") as byte[];
                            writeLineDlgt($"{regKeyPairAppId} setting value AccessPermission = {FormatPermissions(rawAccessPermissions)}");
                            isAccessPermissionSet = rawAccessPermissions?.Length > 0;
                        }
                    }
                    if(!isLaunchPermissionSet)
                    {
                        byte[] rawLaunchPermissions = regKeyPairAppId.GetValue("LaunchPermission") as byte[];
                        RawSecurityDescriptor? sdLaunchPermissions = GetRawSecurityDescriptor(launchRights, launchAsSid, rawLaunchPermissions);
                        writeLineDlgt($"{regKeyPairAppId} current value LaunchPermission = {FormatPermissions(rawLaunchPermissions)}");
                        if(!(isLaunchPermissionSet = sdLaunchPermissions == null))
                        {
                            rawLaunchPermissions = new byte[sdLaunchPermissions.BinaryLength];
                            sdLaunchPermissions.GetBinaryForm(rawLaunchPermissions, 0);
                            regKeyPairAppId.SetValue("LaunchPermission", rawLaunchPermissions);
                            rawLaunchPermissions = regKeyPairAppId.GetValue("LaunchPermission") as byte[];
                            writeLineDlgt($"{regKeyPairAppId} setting value LaunchPermission = {FormatPermissions(rawLaunchPermissions)}");
                            isLaunchPermissionSet = rawLaunchPermissions?.Length > 0;
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Debug.WriteLine(e);
                throw;
            }
            return IsPermissionsSet();
        }
    }
    public class EnvironmentHelper
    {
		public static string GetPublishDirectory(Type type)
		{
			string strPublishDirectory = null;
			Assembly curAssembly = Assembly.GetAssembly(type);
			Assembly entryAssembly = Assembly.GetEntryAssembly();
			FileInfo fileInfoModule = null;
			DirectoryInfo directoryInfo = null;
			if(entryAssembly != null && curAssembly != null && curAssembly != entryAssembly)
			{
				fileInfoModule = new FileInfo(entryAssembly.Location);
				DirectoryInfo assemblyDirectoryInfo = new DirectoryInfo(fileInfoModule.DirectoryName);
				bool found = false;
				while(!found && assemblyDirectoryInfo != null && !string.Equals(assemblyDirectoryInfo.Name, "Visual Studio Projects"))
				{
					found = assemblyDirectoryInfo.GetFiles()?.Any(fileInfo => string.Equals(fileInfo?.Extension, ".sln")) == true;
					if(!found)
					{
						assemblyDirectoryInfo = assemblyDirectoryInfo.Parent;
					}
				}
				if(found)
				{
					directoryInfo = new DirectoryInfo($@"{assemblyDirectoryInfo.FullName}\\{curAssembly.GetName().Name}");
					if(!directoryInfo.Exists)
					{
						directoryInfo = null;
					}
				}
			}
			else
			{
				fileInfoModule = new FileInfo(type.Module.FullyQualifiedName);
				directoryInfo = new DirectoryInfo(fileInfoModule.DirectoryName);
			}
			if(directoryInfo != null)
			{
				FileInfo fileInfoFolderProfile = null;
				FileInfo[] filesInfo = null;
				EnumerationOptions options = new EnumerationOptions()
				{
					MatchCasing = MatchCasing.CaseInsensitive,
					MatchType = MatchType.Simple,
					RecurseSubdirectories = true
				};
				while(fileInfoFolderProfile == null)
				{
					filesInfo = directoryInfo.GetFiles("FolderProfile.pubxml", options);
					if(filesInfo?.Length > 0)
					{
						fileInfoFolderProfile = filesInfo[0];
					}
					else
					{
						directoryInfo = Directory.GetParent(directoryInfo.FullName);
					}
				}
				if(fileInfoFolderProfile != null)
				{
					XmlDocument xmlDocument = new XmlDocument();
					xmlDocument.Load(fileInfoFolderProfile.FullName);
					XmlElement xmlElement = xmlDocument.SelectSingleNode("//PublishDir") as XmlElement;
					strPublishDirectory = xmlElement.InnerText?.Trim();
				}
			}
			return strPublishDirectory;
		}
    }
}