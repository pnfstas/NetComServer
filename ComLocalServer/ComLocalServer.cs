using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

using static NetComServer.Helpers.Win32Helper;
using static NetComServer.Helpers.RegistryHelper;
using static NetComServer.Helpers.EnvironmentHelper;
using static NetComServer.ComServerConstants;
using static NetComServer.ComLocalServer;
using IClassFactory = NetComServer.Helpers.Win32Helper.IClassFactory;
using WriteLineDelegate = NetComServer.Helpers.WriteLineDelegate;

//[assembly: Guid("467E7566-BEF4-4138-A619-035D3171AFCF")]
//[assembly: TypeLibVersion(1, 0)]

return NetComServer.Program.Main(args);

namespace NetComServer
{
	public delegate string? ReadLineDelegate();

	public partial class Program
	{
		[STAThread]
		public static int Main(string[] args)
		{
			int exitCode = -1;
			string[]? arguments = string.Join("", args)?.ToLower()?.Split(['/', '-'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
			Update = arguments?.Contains("update") == true;
			KeepConsole = arguments?.Contains("keepconsole") == true;
			if(arguments?.Contains("unregister") == true)
			{
				exitCode = UnRegisterServer() ? 0 : -1;
			}
			else if(Update || arguments?.Contains("register") == true)
			{
				exitCode = RegisterServer() ? 0 : -1;
			}
			else
			{
				ComLocalServer localServer = new ComLocalServer();
				WriteLine("NetComServer is running...");
				WriteLine($"KeepConsole = {KeepConsole}");
				exitCode = localServer.RegisterClassObject() ? 0 : -1;
				WriteLine("RegisterClassObject is " + (exitCode == 0 ? "succseed." : "failed."));
				/*
				if(exitCode == 0)
				{
					localServer.IsServerRunning = true;
					Thread threadRunServer = new Thread(localServer.Run);
					WriteLine("Message loop is start.");
					MSG msg;
					while(GetMessage(out msg, nint.Zero, 0, 0) > 0)
					{
						TranslateMessage(ref msg);
						DispatchMessage(ref msg);
					}
					WriteLine("Message loop is end.");
					exitCode = localServer.UnRegisterClassObject() ? 0 : -1;
					WriteLine("UnRegisterClassObject is " + (exitCode == 0 ? "succseed." : "failed."));
					WriteLine("Wait for server stop...");
					localServer.IsServerRunning = false;
					localServer.ServerStopEvent.WaitOne();
					localServer.ServerStopEvent.Reset();
					WriteLine("Server is stoped.");
				}
				*/
				if(HasConsole)
				{
					Console.WriteLine("Press any key to exit");
					Console.ReadLine();
					if(!KeepConsole)
					{
						FreeConsole();
					}
				}
			}
			return exitCode;
		}
	}
	[ComVisible(true)]
	[Guid(CLSID_Factory)]
	public class ComServerClassFactory : IClassFactory
	{
		public int QueryInterface([In] ref Guid riid, [Out] out nint ppvObject)
		{
			int hr = S_OK;
			ppvObject = nint.Zero;
			string? strIId = riid.ToString()?.ToUpper()?.Trim();
			switch(strIId)
			{
			case IID_IComServer:
				ppvObject = Marshal.GetComInterfaceForObject<ComServerClassFactory, IComServer>(this);
				hr = Marshal.GetLastSystemError();
				break;
			case IID_IClassFactory:
				ppvObject = Marshal.GetComInterfaceForObject<ComServerClassFactory, IClassFactory>(this);
				hr = Marshal.GetLastSystemError();
				break;
			case IID_IDispatch:
				ppvObject = Marshal.GetIDispatchForObject(this);
				hr = Marshal.GetLastSystemError();
				break;
			case IID_IUnknown:
				ppvObject = Marshal.GetIUnknownForObject(this);
				hr = Marshal.GetLastSystemError();
				break;
			default:
				hr = E_NOINTERFACE;
				break;
			}
			return hr;
		}
		public int CreateInstance([In, MarshalAs(UnmanagedType.IUnknown)] nint pUnkOuter, [In] ref Guid riid, [Out] out nint ppvObject)
		{
			int hr = S_OK;
			ppvObject = nint.Zero;
			if(pUnkOuter == nint.Zero)
			{
				string? strIId = riid.ToString()?.ToUpper()?.Trim();
				switch(strIId)
				{
				case IID_IComServer:
					ppvObject = Marshal.GetComInterfaceForObject<ComServer, IComServer>(new ComServer());
					hr = Marshal.GetLastSystemError();
					break;
				case IID_IDispatch:
					ppvObject = Marshal.GetIDispatchForObject(new ComServer());
					hr = Marshal.GetLastSystemError();
					break;
				case IID_IUnknown:
					ppvObject = Marshal.GetIUnknownForObject(new ComServer());
					hr = Marshal.GetLastSystemError();
					break;
				default:
					hr = E_NOINTERFACE;
					break;
				}
			}
			else
			{
				hr = CLASS_E_NOAGGREGATION;
			}
			return hr;
		}
		public int LockServer([In] bool fLock)
		{
			return S_OK;
		}
	}
	public class ComLocalServer : IDisposable
	{
		public static ReadLineDelegate? ReadLine { get; set; }
		public static WriteLineDelegate? WriteLine { get; set; }
		private static uint ClassId;
		private bool isRunning = false;
		public bool IsServerRunning
		{
			get
			{
				lock(this)
				{
					return isRunning;
				}
			}
			set
			{
				lock(this)
				{
					value = isRunning;
				}
			}
		}
		public static bool Update { get; set; } = false;
		public static bool HasConsole { get; set; } = false;
		public static bool KeepConsole { get; set; } = false;
		private bool disposed = false;
		public ManualResetEvent ServerStopEvent { get; set; } = new ManualResetEvent(false);
		public ComLocalServer()
		{
			HasConsole = HasConsole || ShowConsole(true);
			WriteLine = HasConsole ? Console.WriteLine : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
			ReadLine = HasConsole ? Console.ReadLine : null;
		}
		[UnmanagedCallersOnly(CallConvs = [ typeof(CallConvStdcall) ])]
		public static unsafe int DllGetClassObject(void* rclsid, void* riid, nint* ppv)
		{
			WriteLine = HasConsole ? new WriteLineDelegate(Console.WriteLine) : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
			*ppv = nint.Zero;
			nint ptrClsid = new nint(rclsid);
			nint ptrIId = new nint(riid);
			Guid clsid = Marshal.PtrToStructure<Guid>(ptrClsid);
			Guid iid = Marshal.PtrToStructure<Guid>(ptrIId);
			//WriteLine($"DllGetClassObject: rclsid = {clsid}, riid = {iid}");
			string strMessage = $"DllGetClassObject: rclsid = {clsid}, riid = {iid}";
			int hr = CLASS_E_CLASSNOTAVAILABLE;
			if(CLSID_ComServer.Equals(clsid.ToString().ToUpper().Trim()) && IID_IClassFactory.Equals(iid.ToString().ToUpper().Trim()))
			{
				ComServerClassFactory factory = new ComServerClassFactory();
				nint punk = nint.Zero;
				hr = factory.QueryInterface(ref iid, out punk);
				if(hr != S_OK && punk != nint.Zero)
				{
					*ppv = punk;
				}
			}
			//WriteLine($"DllGetClassObject: hr = {hr:X}");
			strMessage += $"\nDllGetClassObject: hr = {hr:X}";
			MessageBox(nint.Zero, strMessage, "", MessageBoxFlag.MB_OK | MessageBoxFlag.MB_ICONINFORMATION);
			return hr;
		}
		public static bool RegisterServer()
		{
			bool isRegistered = false;
			try
			{
				ComServer.HasConsole = HasConsole = HasConsole || ShowConsole();
				WriteLine = HasConsole ? new WriteLineDelegate(Console.WriteLine) : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
				string strPublishDirectory = GetPublishDirectory(typeof(ComLocalServer));
				string strModuleName = typeof(ComLocalServer).Module.Name.Replace(".dll", ".exe");
				if(!string.IsNullOrWhiteSpace(strModuleName) && !string.IsNullOrWhiteSpace(strPublishDirectory))
				{
					string strServerPath = Path.Combine(strPublishDirectory, strModuleName);
					isRegistered = ComServer.RegisterServer(strServerPath, Update, KeepConsole);
				}
				if(HasConsole = GetConsoleWindow() != nint.Zero)
				{
					if(KeepConsole)
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
		public static bool UnRegisterServer()
		{
			bool isUnregistered = false;
			try
			{
				ComServer.HasConsole = HasConsole = HasConsole || ShowConsole();
				WriteLine = HasConsole ? new WriteLineDelegate(Console.WriteLine) : new WriteLineDelegate((string? message) => Debug.WriteLine(message));
				isUnregistered = ComServer.UnregisterServer(KeepConsole);
				if(HasConsole = GetConsoleWindow() != nint.Zero)
				{
					if(KeepConsole)
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
		public bool RegisterClassObject()
		{
			Guid clsidComServer = new Guid(CLSID_ComServer);
			ComServerClassFactory factory = new ComServerClassFactory();
			bool isRegistered = CoRegisterClassObject(ref clsidComServer, factory, CLSCTX.LOCAL_SERVER, REGCLS.MULTIPLEUSE | REGCLS.SUSPENDED, out ClassId) == S_OK;
			if(isRegistered)
			{
				isRegistered = CoResumeClassObjects() == S_OK;
			}
			return isRegistered;
		}
		public bool UnRegisterClassObject()
		{
			return ClassId == 0 || CoRevokeClassObject(ClassId) == S_OK;
		}
		public void Run()
		{
			while(IsServerRunning)
			{
				GC.Collect();
				Thread.Sleep(2000);
			}
			ServerStopEvent.Set();
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
					UnRegisterClassObject();
					if(HasConsole && !KeepConsole)
					{
						FreeConsole();
					}
				}
				disposed = true;
			}
		}
	}
}