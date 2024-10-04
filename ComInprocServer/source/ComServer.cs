using Microsoft.Win32;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

using static NetComServer.Helpers.Win32Helper;
using static NetComServer.Helpers.RegistryHelper;
using static NetComServer.Helpers.EnvironmentHelper;
using static NetComServer.ComServerConstants;

using WriteLineDelegate = NetComServer.Helpers.WriteLineDelegate;

namespace NetComServer
{
	public enum ComServerType
	{
		Inproc,
		Local
	}
	[ComVisible(true)]
	[Guid("9812DE38-ACE5-426F-9C54-652322461AFC")]
	public class ComServerConstants
	{
		public const string LIBID = "467E7566-BEF4-4138-A619-035D3171AFCF";
		public const string CLSID_Factory = "2A8CABCA-0EBD-40F3-96A4-04054B3ED79D";
		public const string IID_IComServer = "E62A1CB0-86A7-40AE-AFE4-75562C32A498";
		public const string CLSID_ComServer = "C6819A04-046E-4EA2-9750-949760CB26F9";
		public const string ProgIdComServer = "NetComServer.ComServer";
		public const string ProgIdComServerConstants = "NetComServer.ComServerConstants";
		public const string ApplicationName = "NetComServer";
	}
	[ComVisible(true)]
	[Guid(IID_IComServer)]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface IComServer : IDisposable
	{
		abstract int LoadAssembly([In, MarshalAs(UnmanagedType.LPWStr)] string assemblyName);
		[return: MarshalAs(UnmanagedType.IDispatch)]
		abstract object GetCurrentAssembly();
		[return: MarshalAs(UnmanagedType.IDispatch)]
		abstract object CreateNetInstance([In, MarshalAs(UnmanagedType.LPWStr)] string typeName);
		/*
		int LoadAssembly(string assemblyName);
		object GetCurrentAssembly();
		object CreateInstance(string typeName);
		*/
	}
	[ComVisible(true)]
	[Guid(CLSID_ComServer)]
	[ProgId("NetComServer.ComServer")]
	[ClassInterface(ClassInterfaceType.None)]
	[ComDefaultInterface(typeof(IComServer))]
	public class ComServer : IComServer
	{
		private static WriteLineDelegate WriteLine { get; set; }
		public static RegistryKeyPair RegKeyPairRoot { get; } = new RegistryKeyPair(Registry.LocalMachine.OpenSubKey("SOFTWARE\\Classes", true), Registry.ClassesRoot);
		public static string PublishDirectory { get; set; }
		public static bool HasConsole { get; set; } = false;
		public Assembly? CurrentAssembly { get; set; }
		private static ushort VersionMajor { get; set; } = 0;
		private static ushort VersionMinor { get; set; } = 0;
		private static int LCID { get; } = CultureInfo.CurrentCulture.LCID;
		private bool disposed = false;
		public ComServer()
		{
		}
		int IComServer.LoadAssembly(string assemblyName)
		{
			try
			{
				Uri uriAssembly = null;
				CurrentAssembly = Uri.TryCreate(assemblyName, UriKind.Absolute, out uriAssembly) ? Assembly.LoadFrom(assemblyName) : Assembly.Load(assemblyName);
			}
			catch(Exception e)
			{
				CurrentAssembly = null;
				Debug.WriteLine(e);
                throw;
			}
			return CurrentAssembly != null ? S_OK : E_FAIL;
		}
		object IComServer.CreateNetInstance(string typeName)
		{
			object? ret = null;
			try
			{
				ret = CurrentAssembly?.CreateInstance(typeName);
			}
			catch(Exception e)
			{
				ret = null;
				Debug.WriteLine(e);
                throw;
			}
			return ret;
		}
		object IComServer.GetCurrentAssembly()
		{
			return CurrentAssembly;
		}
		public static bool RegisterServer(bool keepConsole = false)
		{
			return RegisterServer(null, false, keepConsole);
		}
		public static bool RegisterServer(bool update, bool keepConsole = false)
		{
			return RegisterServer(null, update, keepConsole);
		}
		public static bool RegisterServer(string strLocalServerPath, bool keepConsole = false)
		{
			return RegisterServer(strLocalServerPath, false, keepConsole);
		}
		public static bool RegisterServer(string? strLocalServerPath, bool update, bool keepConsole = false)
		{
			bool isRegistered = false;
			try
			{
				HasConsole = HasConsole || ShowConsole();
				WriteLine = HasConsole ? new WriteLineDelegate(Console.WriteLine) : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
				Guid clsid = new Guid(CLSID_ComServer);
				Guid libid = new Guid(LIBID);
				string clsidKeyName = string.Format(@"CLSID\{0:B}", clsid);
				string appIdKeyName = string.Format(@"AppID\{0:B}", clsid);
				RegistryKeyPair regKeyPairClsid = RegKeyPairRoot.OpenSubKey(clsidKeyName, true);
				RegistryKeyPair regKeyPairAppId = RegKeyPairRoot.OpenSubKey(appIdKeyName, true);
				RegistryKeyPair regKeyPairProgId = null;
				ComServerType serverType = !string.IsNullOrWhiteSpace(strLocalServerPath) ? ComServerType.Local : ComServerType.Inproc;
				string strServerPath = null;
				if(serverType == ComServerType.Local)
				{
					strServerPath = strLocalServerPath;
					isRegistered = !update && regKeyPairClsid != null && regKeyPairAppId != null && IsRegisteredTypeLibrary();
				}
				else
				{
					Type type = typeof(ComServer);
					string strModuleName = type.Module?.Name?.Trim()?.Replace(".dll", ".comhost.dll");
					if(!string.IsNullOrWhiteSpace(strModuleName))
					{
						PublishDirectory = GetPublishDirectory(type);
						if(!string.IsNullOrWhiteSpace(PublishDirectory))
						{
							strServerPath = Path.Combine(PublishDirectory, strModuleName);
						}
					}
					regKeyPairProgId = RegKeyPairRoot.OpenSubKey(ProgIdComServer, true);
					isRegistered = !update && regKeyPairClsid != null && regKeyPairAppId != null && regKeyPairProgId != null && IsRegisteredTypeLibrary();
				}
				if(!isRegistered && !string.IsNullOrWhiteSpace(strServerPath) && RegisterTypeLibrary(update))
				{
					if(regKeyPairClsid != null)
					{
						WriteLine($"{regKeyPairClsid} and subtree were deleted.");
						RegKeyPairRoot.DeleteSubKeyTree(clsidKeyName, false);
					}
					if(regKeyPairAppId != null)
					{
						WriteLine($"{regKeyPairAppId} and subtree were deleted.");
						RegKeyPairRoot.DeleteSubKeyTree(appIdKeyName, false);
					}
					if(regKeyPairProgId != null)
					{
						WriteLine($"{regKeyPairProgId} and subtree were deleted.");
						RegKeyPairRoot.DeleteSubKeyTree(ProgIdComServer, false);
					}
					VersionMajor = Math.Max(VersionMajor, (ushort)1);
					VersionMinor = Math.Max(VersionMinor, (ushort)0);
					// HKLM\SOFTWARE\Classes\CLSID
					regKeyPairClsid = RegKeyPairRoot.CreateSubKey(clsidKeyName);
					WriteLine(regKeyPairClsid + (regKeyPairClsid.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPairClsid.SetValue(null, ProgIdComServer);
					WriteLine($"{regKeyPairClsid} setting default value = {regKeyPairClsid.GetValue(null)}");
					regKeyPairClsid.SetValue("AppID", clsid.ToString("B"));
					WriteLine($"{regKeyPairClsid} setting value AppID = {regKeyPairClsid.GetValue("AppID")}");
					RegistryKeyPair regKeyPair = null;
					if(serverType == ComServerType.Inproc)
					{
						regKeyPair = regKeyPairClsid.CreateSubKey("InProcServer32");
						WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
						regKeyPair.SetValue(null, strServerPath);
						WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
						regKeyPair.SetValue("ThreadingModel", "Both");
						WriteLine($"{regKeyPair} setting value ThreadingModel = {regKeyPair.GetValue("ThreadingModel")}");
					}
					else
					{
						regKeyPair = regKeyPairClsid.CreateSubKey("LocalServer32");
						WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
						regKeyPair.SetValue(null, strLocalServerPath.Trim());
						WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
					}
					regKeyPair = regKeyPairClsid.CreateSubKey("ProgId");
					WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPair.SetValue(null, $"{ProgIdComServer}.{VersionMajor}");
					regKeyPair.SetValue(null, ProgIdComServer);
					WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
					/*
					regKeyPair = regKeyPairClsid.CreateSubKey("VersionIndependentProgID");
					WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPair.SetValue(null, ProgIdComServer);
					WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
					*/
					regKeyPair = regKeyPairClsid.CreateSubKey("Version");
					WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPair.SetValue(null, $"{VersionMajor}.{VersionMinor}");
					WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
					regKeyPair = regKeyPairClsid.CreateSubKey("TypeLib");
					WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPair.SetValue(null, libid.ToString("B"));
					WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
					// HKLM\SOFTWARE\Classes\{ProgIdComServer}
					regKeyPairProgId = RegKeyPairRoot.CreateSubKey(ProgIdComServer);
					WriteLine(regKeyPairProgId + (regKeyPairProgId.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPairProgId.SetValue(null, ProgIdComServer);
					WriteLine($"{regKeyPairProgId} setting default value = {regKeyPairProgId.GetValue(null)}");
					regKeyPair = regKeyPairProgId.CreateSubKey("CLSID");
					WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPair.SetValue(null, clsid.ToString("B"));
					WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
					// HKLM\SOFTWARE\Classes\AppID\{CLSID_ComServer}
					regKeyPairAppId = RegKeyPairRoot.CreateSubKey(appIdKeyName);
					WriteLine(regKeyPairAppId + (regKeyPairAppId.KeyCount > 1 ? "were" : "was") + " created");
					regKeyPairAppId.SetValue(null, ApplicationName);
					WriteLine($"{regKeyPairAppId} setting default value = {regKeyPairAppId.GetValue(null)}");
					/*
					regKeyPairAppId.SetValue("DllSurrogate", string.Empty);
					WriteLine($"{regKeyPairAppId} setting value DllSurrogate = {regKeyPairAppId.GetValue("DllSurrogate")}");
					regKeyPairAppId.SetValue("PreferredServerBitness", 3, RegistryValueKind.DWord);
					WriteLine($"{regKeyPairAppId} setting value PreferredServerBitness = {regKeyPairAppId.GetValue("PreferredServerBitness")}");
					*/
					COM_RIGHTS accessRights = COM_RIGHTS.EXECUTE | COM_RIGHTS.EXECUTE_LOCAL;
					COM_RIGHTS launchRights = COM_RIGHTS.EXECUTE | COM_RIGHTS.EXECUTE_LOCAL | COM_RIGHTS.ACTIVATE_LOCAL;
					string accountName = "BUILTIN\\Пользователи";
					Helpers.SecurityHelper.SetComServerPermissions(regKeyPairAppId, accessRights, launchRights, accountName, accountName, WriteLine);
					accountName = "NT AUTHORITY\\LocalService";
					Helpers.SecurityHelper.SetComServerPermissions(regKeyPairAppId, accessRights, launchRights, accountName, accountName, WriteLine);
					accountName = "ИНТЕРАКТИВНЫЕ";
					Helpers.SecurityHelper.SetComServerPermissions(regKeyPairAppId, accessRights, launchRights, accountName, accountName, WriteLine);
					regKeyPairAppId.SetValue("RunAs", "Interactive User");
					WriteLine($"{regKeyPairAppId} setting value RunAs = {regKeyPairAppId.GetValue("RunAs")}");
					isRegistered = regKeyPairClsid != null && regKeyPairAppId != null && regKeyPairProgId != null;
				}
				if(HasConsole)
				{
					if(keepConsole)
					{
						Console.ReadLine();
					}
					else
					{
						FreeConsole();
					}
				}
			}
			catch(Exception e)
			{
				isRegistered = false;
				Debug.WriteLine(e);
				throw;
			}
			return isRegistered;
		}
		public static bool UnregisterServer(bool keepConsole = false)
		{
			bool isUnregistered = false;
			try
			{
				HasConsole = HasConsole || ShowConsole();
				WriteLine = HasConsole ? new WriteLineDelegate(Console.WriteLine) : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
				Type typeServer = typeof(ComServer);
				Type typeGuids = typeof(ComServerConstants);
				AttributeCollection comServerAttributes = TypeDescriptor.GetAttributes(typeServer);
				AttributeCollection comServerGuidsAttributes = TypeDescriptor.GetAttributes(typeGuids);
				GuidAttribute comServerGuidsAttribute = comServerAttributes[typeof(GuidAttribute)] as GuidAttribute;
				ProgIdAttribute comServerProgIdAttribute = comServerAttributes[typeof(ProgIdAttribute)] as ProgIdAttribute;
				ProgIdAttribute comServerGuidsProgIdAttribute = comServerGuidsAttributes[typeof(ProgIdAttribute)] as ProgIdAttribute;
				string progIdServer = comServerProgIdAttribute?.Value ?? typeServer.FullName;
				string progIdGuids = comServerGuidsProgIdAttribute?.Value ?? typeGuids.FullName;
				Guid clsid = new Guid(comServerGuidsAttribute.Value);
				string clsidKeyName = string.Format(@"CLSID\{0:B}", clsid);
				string appIdKeyName = string.Format(@"AppID\{0:B}", clsid);
				RegistryKeyPair regKeyPair = RegKeyPairRoot.OpenSubKey(clsidKeyName, true);
				if(regKeyPair != null)
				{
					WriteLine($"{regKeyPair} and subtree were deleted.");
					RegKeyPairRoot.DeleteSubKeyTree(clsidKeyName, false);
				}
				regKeyPair = RegKeyPairRoot.OpenSubKey(appIdKeyName, true);
				if(regKeyPair != null)
				{
					WriteLine($"{regKeyPair} and subtree were deleted.");
					RegKeyPairRoot.DeleteSubKeyTree(appIdKeyName, false);
				}
				regKeyPair = RegKeyPairRoot.OpenSubKey(ProgIdComServer, true);
				if(regKeyPair != null)
				{
					WriteLine($"{regKeyPair} and subtree were deleted.");
					RegKeyPairRoot.DeleteSubKeyTree(ProgIdComServer, false);
				}
				regKeyPair = RegKeyPairRoot.OpenSubKey(ProgIdComServerConstants, true);
				if(regKeyPair != null)
				{
					WriteLine($"{regKeyPair} and subtree were deleted.");
					RegKeyPairRoot.DeleteSubKeyTree(ProgIdComServerConstants, false);
				}
				isUnregistered = RegKeyPairRoot.OpenSubKey(clsidKeyName) == null
					&& RegKeyPairRoot.OpenSubKey(appIdKeyName) == null
					&& RegKeyPairRoot.OpenSubKey(ProgIdComServer) == null
					&& RegKeyPairRoot.OpenSubKey(ProgIdComServerConstants) == null
					&& UnregisterTypeLibrary();
				if(HasConsole)
				{
					if(keepConsole)
					{
						Console.ReadLine();
					}
					else
					{
						FreeConsole();
					}
				}
			}
			catch(Exception e)
			{
				isUnregistered = false;
				Debug.WriteLine(e);
				throw;
			}
			return isUnregistered;
		}
		public static bool IsRegisteredTypeLibrary()
		{
			bool isRegistered = false;
			try
			{
				Guid libid = new Guid(LIBID);
				Guid riid = new Guid(IID_IComServer);
				string tlbKeyName = string.Format(@"TypeLib\{0:B}", libid);
				string interfaceKeyName = string.Format(@"Interface\{0:B}", riid);
				RegistryKeyPair regKeyPairTypeLib = RegKeyPairRoot.OpenSubKey(tlbKeyName);
				RegistryKeyPair regKeyPairInterface = RegKeyPairRoot.OpenSubKey(interfaceKeyName);
				isRegistered = regKeyPairTypeLib != null && regKeyPairInterface != null;
			}
			catch(Exception e)
			{
				isRegistered = false;
				Debug.WriteLine(e);
				throw;
			}
			return isRegistered;
		}
		public static bool RegisterTypeLibrary(bool update = false)
		{
			bool isRegistered = false;
			try
			{
				WriteLine = HasConsole ? new WriteLineDelegate(Console.WriteLine) : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
				Guid libid = new Guid(LIBID);
				Guid riid = new Guid(IID_IComServer);
				string tlbKeyName = string.Format(@"TypeLib\{0:B}", libid);
				string interfaceKeyName = string.Format(@"Interface\{0:B}", riid);
				RegistryKeyPair regKeyPairTypeLib = RegKeyPairRoot.OpenSubKey(tlbKeyName);
				RegistryKeyPair regKeyPairInterface = RegKeyPairRoot.OpenSubKey(interfaceKeyName);
				isRegistered = !update && regKeyPairTypeLib != null && regKeyPairInterface != null;
				if(!isRegistered)
				{
					if(string.IsNullOrWhiteSpace(PublishDirectory))
					{
						PublishDirectory = GetPublishDirectory(typeof(ComServer));
					}
					if(!string.IsNullOrWhiteSpace(PublishDirectory))
					{
						if(UnregisterTypeLibrary())
						{
							string tlbPath = Path.Combine(PublishDirectory, "lib\\NetComServer.tlb");
							ITypeLib typeLib = null;
							int hr = LoadTypeLibEx(tlbPath, REGKIND.NONE, out typeLib);
							WriteLine($"Register type library with hr = 0x{string.Format("{0:X}", hr)}");
							if(hr == S_OK)
							{
								Guid guidOleAut32 = default;
								RegistryKey regKeyClsid = Registry.ClassesRoot.OpenSubKey("CLSID");
								string[] arrClsids = regKeyClsid.GetSubKeyNames();
								RegistryKey regKey = null;
								string regValue = null;
								foreach(string clsid in arrClsids)
								{
									regKey = regKeyClsid.OpenSubKey(clsid);
									regValue = regKey?.GetValue(null) as string;
									if(regValue != null && regValue.Trim().Equals("PSOAInterface"))
									{
										guidOleAut32 = new Guid(clsid);
										break;
									}
								}
								if(guidOleAut32 != default)
								{
									nint pTlibAttr = nint.Zero;
									typeLib.GetLibAttr(out pTlibAttr);
									if(pTlibAttr != nint.Zero)
									{
										TYPELIBATTR tlibAttr = Marshal.PtrToStructure<TYPELIBATTR>(pTlibAttr);
										string tlibName = null;
										string tlibDescription = null;
										string tlibHelpFile = null;
										typeLib.GetDocumentation(-1, out tlibName, out tlibDescription, out int _, out tlibHelpFile);
										tlibDescription = !string.IsNullOrWhiteSpace(tlibDescription) ? tlibDescription :
											!string.IsNullOrWhiteSpace(tlibName) ? tlibName : typeof(ComServer).Namespace;
										string tlibVersion = $"{tlibAttr.wMajorVerNum}.{tlibAttr.wMinorVerNum}";
										string tlibSysKind = tlibAttr.syskind == SYSKIND.SYS_WIN64 ? "win64" : "win32";
										string tlibHelpDir = PublishDirectory;
										if(!string.IsNullOrWhiteSpace(tlibHelpFile))
										{
											FileInfo fileInfo = new FileInfo(tlibHelpFile);
											if(!string.IsNullOrWhiteSpace(fileInfo.DirectoryName))
											{
												tlibHelpDir = fileInfo.DirectoryName;
											}
										}
										//TypeLib
										regKeyPairTypeLib = RegKeyPairRoot.CreateSubKey(tlbKeyName);
										WriteLine(regKeyPairTypeLib + (regKeyPairTypeLib.KeyCount > 1 ? "were" : "was") + " created");
										RegistryKeyPair regKeyPairVersion = regKeyPairTypeLib.CreateSubKey(tlibVersion);
										WriteLine(regKeyPairVersion + (regKeyPairVersion.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPairVersion.SetValue(null, tlibDescription);
										WriteLine($"{regKeyPairVersion} setting default value = {regKeyPairVersion.GetValue(null)}");
										RegistryKeyPair regKeyPair = regKeyPairVersion.CreateSubKey(tlibAttr.lcid.ToString("D"));
										WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPair = regKeyPair.CreateSubKey(tlibSysKind);
										WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPair.SetValue(null, tlbPath);
										WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
										regKeyPair = regKeyPairVersion.CreateSubKey("FLAGS");
										WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPair.SetValue(null, tlibAttr.wLibFlags.ToString("D"));
										WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
										regKeyPair = regKeyPairVersion.CreateSubKey("HELPDIR");
										WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPair.SetValue(null, tlibHelpDir);
										WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
										//Interface
										regKeyPairInterface = RegKeyPairRoot.CreateSubKey(interfaceKeyName);
										WriteLine(regKeyPairInterface + (regKeyPairInterface.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPairInterface.SetValue(null, nameof(IComServer));
										WriteLine($"{regKeyPairInterface} setting default value = {regKeyPairInterface.GetValue(null)}");
										regKeyPair = regKeyPairInterface.CreateSubKey("ProxyStubClsid32");
										WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPair.SetValue(null, guidOleAut32.ToString("B"));
										WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
										regKeyPair = regKeyPairInterface.CreateSubKey("TypeLib");
										WriteLine(regKeyPair + (regKeyPair.KeyCount > 1 ? "were" : "was") + " created");
										regKeyPair.SetValue(null, libid.ToString("B"));
										WriteLine($"{regKeyPair} setting default value = {regKeyPair.GetValue(null)}");
										regKeyPair.SetValue("Version", tlibVersion);
										WriteLine($"{regKeyPair} setting value Version = {regKeyPair.GetValue("Version")}");
									}
								}
							}
						}
					}
					isRegistered = regKeyPairTypeLib != null && regKeyPairInterface != null;
				}
			}
			catch(Exception e)
			{
				isRegistered = false;
				Debug.WriteLine(e);
				throw;
			}
			return isRegistered;
		}
		public static bool UnregisterTypeLibrary()
		{
			bool isUnregistered = false;
			try
			{
				WriteLine = HasConsole ? new WriteLineDelegate(Console.WriteLine) : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
				Guid libid = new Guid(LIBID);
				Guid riid = new Guid(IID_IComServer);
				string tlbKeyName = string.Format(@"TypeLib\{0:B}", libid);
				string interfaceKeyName = string.Format(@"Interface\{0:B}", riid);
				RegistryKeyPair regKeyPairTypeLib = RegKeyPairRoot.OpenSubKey(tlbKeyName, true);
				RegistryKeyPair regKeyPairInterface = RegKeyPairRoot.OpenSubKey(interfaceKeyName, true);
				if(regKeyPairTypeLib != null || regKeyPairInterface != null)
				{
					if(regKeyPairTypeLib != null)
					{
						WriteLine($"{regKeyPairTypeLib} and subtree were deleted.");
						RegKeyPairRoot.DeleteSubKeyTree(tlbKeyName, false);
					}
					if(regKeyPairInterface != null)
					{
						WriteLine($"{regKeyPairInterface} and subtree were deleted.");
						RegKeyPairRoot.DeleteSubKeyTree(interfaceKeyName, false);
					}
				}
				isUnregistered = RegKeyPairRoot.OpenSubKey(tlbKeyName) == null
					&& RegKeyPairRoot.OpenSubKey(interfaceKeyName) == null;
			}
			catch(Exception e)
			{
				isUnregistered = false;
				Debug.WriteLine(e);
				throw;
			}
			return isUnregistered;
		}
		[ComRegisterFunction]
		public static void RegisterFunction(Type type)
		{
			if(type == typeof(ComServer))
			{
				RegisterServer(true, false);
			}
		}
		[ComUnregisterFunction]
		public static void UnregisterFunction(Type type)
		{
			if(type == typeof(ComServer))
			{
				UnregisterServer();
			}
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
					GC.Collect();
					GC.WaitForPendingFinalizers();
				}
				disposed = true;
			}
		}
	}
}
