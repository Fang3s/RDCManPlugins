using RdcMan;
using System.Windows.Forms;
using System.Xml;
using System.ComponentModel.Composition;
using System.Diagnostics;
using static RdcMan.RdpClient;
using System.Reflection;
using System;
using System.Runtime.InteropServices;
using MSTSCLib;
using System.Security.Permissions;
using System.Linq;
using System.Windows.Threading;
using System.Threading;
using System.Collections.Generic;
using Timer = System.Windows.Forms.Timer;
/*
Form
+-------------------------------+
|+-----------------------------+|
|| ToolBar                     ||
|+-----------------------------+|
|+-----------------------------+|
||             Body            ||
||+------+|+------------------+||
||| Tree |||                  |||
|||      |||                  |||
|||      |||    Right Panel   |||
|||      |||                  |||
|||      |||                  |||
||+------+v+------------------+||
||  VerticalSplit              ||
|+-----------------------------+|
+-------------------------------+

Window 00130866 "sap-ts03 - Remote Desktop Connection Manager v2.7" WindowsForms10.Window.8.app.0.3d90434_r8_ad1	Form
  Window 00070D5C "" WindowsForms10.Window.8.app.0.3d90434_r8_ad1 (TID=000028F8)									ToolBar
  Window 000D0D82 "" WindowsForms10.Window.8.app.0.3d90434_r8_ad1 (TID=000028F8)
  Window 002F0B82 "" WindowsForms10.Window.8.app.0.3d90434_r8_ad1 (TID=000028F8)									Body
  Window 00200CE6 "" WindowsForms10.Window.8.app.0.3d90434_r8_ad1 (TID=000028F8)									Right Panel
  Window 000B0E72 "" WindowsForms10.SCROLLBAR.app.0.3d90434_r8_ad1 (TID=000028F8)
  Window 00090DD2 "" ATL:00007FFD7958F2F0 (TID=000028F8)
  Window 0019073A "" ATL:00007FFD7958F2F0 (TID=000028F8)									RdpClient.Control (Undockable Window)
	Window 00070D4E "" UIMainClass (TID=000028F8)
	  Window 00070D52 "" UIContainerClass (TID=000028F8)
		Window 00100CD2 "Input Capture Window" IHWindowClass (TID=00005AF0)*
		Window 000C0870 "Output Painter Window" OPContainerClass (TID=000060A0)*
		  Window 00110846 "Output Painter Child Window" OPWindowClass (TID=000060A0)
			Window 001C0956 "Output Painter DX Child Window" OPWindowClass (TID=000060A0)
			Window 00240AE8 "Connected to sap-ts03" WindowsForms10.STATIC.app.0.3d90434_r8_ad1 (TID=000060A0)
  Window 00090766 "" WindowsForms10.Window.8.app.0.3d90434_r8_ad1 (TID=000028F8)									VerticalSplit
  Window 001C0960 "" WindowsForms10.SysTreeView32.app.0.3d90434_r8_ad1 (TID=000028F8)								TreeView


class RdcMan.Program {
	private static PluginContext PluginContext;

	private static void InstantiatePlugins() {
		PluginContext = new PluginContext();
		Assembly callingAssembly = Assembly.GetCallingAssembly();
		DirectoryCatalog catalog = new DirectoryCatalog(Path.GetDirectoryName(callingAssembly.Location), "Plugin.*.dll");
		//...
		item2.PreLoad(PluginContext, value4.SettingsNode);
		//...
	}

	private static void CompleteInitialization() {
		InstantiatePlugins();
		//...
		bool isFirstConnection = ReconnectAtStartup(connectedServers);
		if (_serversToConnect != null) {
			ConnectNamedServers(_serversToConnect, isFirstConnection);
		}
		//...
		PluginAction(delegate(IPlugin p) { p.PostLoad(PluginContext); });
	}
}

Thanks to:
- https://github.com/Tadas/RDCManPlugins.git
- https://github.com/icsharpcode/ILSpy/releases
See also:
- https://github.com/lzpong/RDCMan.git
- https://github.com/Geosong/RdcManPluginFix.git
*/
namespace Plugin.KeepAlive
{
	//https://github.com/mRemoteNG/mRemoteNG/issues/405
	//Simulate keyboard or mouse move to avoid RDP session idle timeouts.
	[Export(typeof(IPlugin))]
	[SecurityPermission(SecurityAction.LinkDemand, Flags = SecurityPermissionFlag.UnmanagedCode)]
	public class PluginKeepAlive : IPlugin, IMessageFilter
	{
		private const int timerInterval = (60 + 58) * 1000; // 1 minute 58 seconds

		//WindowsForms10.SysTreeView32.app.0.3d90434_r8_ad1
		//internal class RdcMan.ServerTree : System.Windows.Forms.TreeView, RdcMan.IServerTree {}
		private IServerTree serverTree;
		private Dictionary<Server, Timer> serverTimers = new Dictionary<Server, Timer>();

		/// <summary>
		/// called when the user right clicks a server node in the tree
		/// </summary>
		/// <param name="contextMenuStrip"></param>
		/// <param name="node"></param>
		public void OnContextMenu(ContextMenuStrip contextMenuStrip, RdcTreeNode node)
		{
			//MessageBox.Show("OnContextMenu", "Plugin.KeepAlive event", MessageBoxButtons.OK, MessageBoxIcon.Information);
			if (null == node as GroupBase)
			{
				if (null != node as ServerBase)
				{
					string targetHost = (node as ServerBase).ServerName;
					ToolStripMenuItem NewMenuItem = new DelegateMenuItem("Enter-PSSession", MenuNames.None, () => this.EnterPSSession(targetHost));
					//NewMenuItem.Image = Properties.Resources.PowerShell5_32;
					contextMenuStrip.Items.Insert(contextMenuStrip.Items.Count - 1, NewMenuItem);
					contextMenuStrip.Items.Insert(contextMenuStrip.Items.Count - 1, new ToolStripSeparator());
				}
			}
		}

		public void OnDockServer(ServerBase server)
		{
			//MessageBox.Show("OnDockServer", "Plugin.KeepAlive event", MessageBoxButtons.OK, MessageBoxIcon.Information);
			Debug.WriteLine("OnDockServer|Plugin.KeepAlive event");
		}

		public void OnUndockServer(IUndockedServerForm form)
		{
			MessageBox.Show("OnUndockServer", "Plugin.KeepAlive event", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		/// <summary>
		/// called after plugins and the connection tree is loaded
		/// </summary>
		/// <param name="context"></param>
		public void PostLoad(IPluginContext context)
		{
			//MessageBox.Show("PostLoad", "Plugin.KeepAlive event", MessageBoxButtons.OK, MessageBoxIcon.Information);
			try
			{
				serverTree = context.Tree;
				(serverTree as TreeView).NodeMouseDoubleClick += PluginKeepAlive_NodeMouseDoubleClick;
				GroupBase rootNode = context.Tree.RootNode;
				Console.WriteLine("RootNode.Password: {}", context.Tree.RootNode.Password);
				Debug.WriteLine(context.Tree.RootNode.Password, "RootNode.Password");
			}
			catch (System.Exception ex)
			{
				Debug.WriteLine(ex.ToString());
			}
		}

		/// <summary>
		/// called while the plugins are loading 此时界面控件还未显示
		/// </summary>
		/// <param name="context"></param>
		/// <param name="xmlNode"></param>
		public void PreLoad(IPluginContext context, XmlNode xmlNode)
		{
			//SetThreadExecutionState(EXECUTION_STATE.ES_MY_NO_SLEEP);
			//MessageBox.Show("PreLoad", "Plugin.KeepAlive event", MessageBoxButtons.OK ,MessageBoxIcon.Information);

			Debug.WriteLine("PreLoad|Plugin.KeepAlive event", Util.C_SendKeys);
			Server.ConnectionStateChanged += Server_ConnectionStateChanged;
			Application.AddMessageFilter(this);
		}

		/// <summary>
		/// user clicked OK in the Options dialog
		/// </summary>
		/// <returns></returns>
		public XmlNode SaveSettings()
		{
			//MessageBox.Show("SaveSettings", "Plugin.KeepAlive event", MessageBoxButtons.OK, MessageBoxIcon.Information);
			Debug.WriteLine("SaveSettings|Plugin.KeepAlive event");
			return null;
		}

		/// <summary>
		/// RDCMan is shutting down
		/// </summary>
		public void Shutdown()
		{
			//MessageBox.Show("Shutdown", "Plugin.KeepAlive event", MessageBoxButtons.OK, MessageBoxIcon.Information);
			Debug.WriteLine("Shutdown|Plugin.KeepAlive event");
		}

		private void EnterPSSession(string targetHost)
		{
			string PSpath = System.IO.Path.Combine(System.Environment.SystemDirectory, "WindowsPowerShell\\v1.0\\powershell.exe");
			string Arguments = "-NoLogo -NoExit Enter-PSSession " + targetHost;
			//MessageBox.Show(PSpath + " " + Arguments, "Starting PSSession", MessageBoxButtons.OK, MessageBoxIcon.Information);
			Process.Start(PSpath, Arguments);
		}

		private void PluginKeepAlive_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
		{
			Debug.WriteLine(e.Node.Text, "NodeMouseDoubleClick");  //throw new System.NotImplementedException();
			Server rdcServer = e.Node as Server;
		}


		[DllImport("Kernel32")]
		public static extern uint GetCurrentThreadId();

		public enum EXECUTION_STATE : uint
		{
			ES_AWAYMODE_REQUIRED = 0x00000040,
			ES_CONTINUOUS = 0x80000000,
			ES_DISPLAY_REQUIRED = 0x00000002,
			ES_SYSTEM_REQUIRED = 0x00000001,
			ES_USER_PRESENT = 0x00000004,
			ES_MY_NO_SLEEP = 0x80000003
		}

		/// <summary>
		/// Enables an application to inform the system that it is in use, thereby preventing the system
		/// from entering sleep or turning off the display while the application is running.
		/// </summary>
		/// <param name="esFlags">The thread's execution requirements.</param>
		/// <returns>
		/// If the function succeeds, the return value is the previous thread execution state.
		/// If the function fails, the return value is NULL.
		/// </returns>
		/// <remarks>https://docs.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadexecutionstate</remarks>
		[DllImport("Kernel32.dll", CharSet = CharSet.Auto)]
		public static extern EXECUTION_STATE SetThreadExecutionState(EXECUTION_STATE esFlags);


		//the drawback is to remain focus to the current window
		//https://stackoverflow.com/questions/17525377/how-to-use-inputsimulator-to-simulate-specific-keys-on-remote-desktop
		[DllImport("user32.dll", EntryPoint = "keybd_event", CharSet = CharSet.Auto,
		ExactSpelling = true)]
		public static extern void keybd_event(byte vk, byte scan, int flags, int extrainfo);

		[DllImport("user32.dll")]
		public static extern IntPtr GetKeyboardLayout(uint idThread);

		[DllImport("user32.dll")]
		public static extern short VkKeyScanEx(char ch, IntPtr dwhkl);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern int MapVirtualKey(int uCode, int uMapType);

		// Marshalling a NULL pointer: If it returns a NULL pointer, you'll get IntPtr.Zero on the managed side.
		// https://docs.microsoft.com/en-us/archive/msdn-magazine/2003/july/net-column-calling-win32-dlls-in-csharp-with-p-invoke

		/// <summary>
		/// Retrieves a handle to the foreground window (the window with which the user is currently working).
		/// The system assigns a slightly higher priority to the thread that creates the foreground window than
		/// it does to other threads.
		/// </summary>
		/// <returns>
		/// The return value is a handle to the foreground window. The foreground window can be NULL
		/// in certain circumstances, such as when a window is losing activation.
		/// </returns>
		[DllImport("user32.dll")]
		private static extern IntPtr GetForegroundWindow();

		/// <summary>
		/// Retrieves the window handle to the active window attached to the calling thread's message queue.
		/// </summary>
		/// <returns>
		/// The return value is the handle to the active window attached to the calling thread's
		/// message queue. Otherwise, the return value is NULL.
		/// </returns>
		/// <remarks>
		/// To get the handle to the foreground window, you can use GetForegroundWindow.
		/// To get the window handle to the active window in the message queue for another thread, use GetGUIThreadInfo.
		/// </remarks>
		[DllImport("user32.dll")]
		private static extern IntPtr GetActiveWindow();

		/// <summary>
		/// Activates a window. The window must be attached to the calling thread's message queue.
		/// </summary>
		/// <param name="hWnd">A handle to the top-level window to be activated.</param>
		/// <returns>
		/// If the function succeeds, the return value is the handle to the window that was previously active.
		/// If the function fails, the return value is NULL. To get extended error information, call GetLastError.
		/// </returns>
		/// <remarks>
		/// - The SetActiveWindow function activates a window, but not if the application is in the
		/// background. The window will be brought into the foreground (top of Z-Order) if its application
		/// is in the foreground when the system activates the window.
		/// - If the window identified by the hWnd parameter was created by the calling thread, the active
		/// window status of the calling thread is set to hWnd.Otherwise, the active window status of the
		/// calling thread is set to NULL.
		/// </remarks>
		/// <see cref="https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setactivewindow"/>
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern IntPtr SetActiveWindow(IntPtr hWnd);

		//https://docs.microsoft.com/en-us/windows/win32/inputdev/wm-activate
		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		public static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, uint lParam);

		[DllImport("user32.dll", SetLastError = true)]
		static extern bool PostMessage(HandleRef hWnd, uint Msg, IntPtr wParam, IntPtr lParam);


		[DllImport("user32.dll", ExactSpelling = true)]
		static extern IntPtr SetTimer(IntPtr hWnd, IntPtr nIDEvent, uint uElapse, TimerProc lpTimerFunc);
		//https://stackoverflow.com/questions/26179691/callbacks-from-c-to-c-sharp
		[UnmanagedFunctionPointer(CallingConvention.StdCall)]
		delegate void TimerProc(IntPtr hWnd, uint uMsg, IntPtr nIDEvent, uint dwTime);
#if false // or alternatively
		[DllImport("user32.dll", ExactSpelling = true)]
		static extern IntPtr SetTimer(IntPtr hWnd, IntPtr nIDEvent, uint uElapse, IntPtr lpTimerFunc);
#endif

		private void Server_ConnectionStateChanged(ConnectionStateChangedEventArgs obj)
		{
			Debug.WriteLine("{0}|{1:x}|Server_ConnectionStateChanged|{2}:{3}", DateTime.Now.ToString("o"), GetCurrentThreadId(), obj.Server.Text, obj.State.ToString());
			RdpClient rdpClient = Util.GetRdpClient(obj.Server);
			if (ConnectionState.Connected == obj.State)
			{
				if (rdpClient.Control.InvokeRequired)
				{
					Debug.WriteLine("in non-UI Thread");
					rdpClient.Control.Invoke(new MethodInvoker(() => StartKeepTimer(obj.Server)));
				}
				else
				{
					StartKeepTimer(obj.Server);
				}
				if (Dispatcher.CurrentDispatcher.Thread == Thread.CurrentThread)
				{
					Debug.WriteLine("Do something on current dispatcher");
				}
				Dispatcher dispatcher = Dispatcher.FromThread(Thread.CurrentThread);
				if (null != dispatcher)
				{
					Debug.WriteLine("We know the thread have a dispatcher that we can use");
				}
			}
			else
			{
				if (serverTimers.TryGetValue(obj.Server, out Timer timer))
				{
					timer.Stop();
					serverTimers.Remove(obj.Server);
				}
			}
			//if Server#ConnectionStateChanged is Disconnected, rdpClient is null.
			if (null != rdpClient)
			{
				Debug.WriteLine(rdpClient.DesktopSize.ToFormattedString(), "Server_ConnectionStateChanged");
				//IMsRdpClientNonScriptable ocx = (IMsRdpClientNonScriptable)rdpClient.MsRdpClient8.GetOcx();
				////https://docs.microsoft.com/zh-cn/windows/win32/termserv/imsrdpclientadvancedsettings-keepaliveinterval
				//rdpClient.MsRdpClient8.AdvancedSettings8.keepAliveInterval = 1;
				//rdpClient.MsRdpClient.AdvancedSettings.allowBackgroundInput = 1;
				////https://docs.microsoft.com/zh-cn/windows/win32/termserv/imsrdpclientadvancedsettings-minutestoidletimeout
				//rdpClient.MsRdpClient8.AdvancedSettings2.MinutesToIdleTimeout = 1;

				//https://docs.microsoft.com/zh-cn/windows/win32/termserv/imsrdpclientnonscriptable-sendkeys
				//Now reference the dll and you can use it with bool[] and int[] parameters. For CTRL+ALT+DEL:
				//0x1d: CONTROL, 0x38: MENU/ALT, 0x53: DEL
				bool[] bools = { false, false, false, true, true, true, };
				int[] ints = { 0x1d, 0x38, 0x53, 0x53, 0x38, 0x1d };
				//Focus helps to avoid most E_FAIL exceptions
				//rdp.Focus();

				//unsafe void SendKeys(int numKeys, int* pbArrayKeyUp, int* plKeyData);
				//fixed (bool* pKeyReleased = bools)
				//fixed (int* pScanCodes = ints)
				//{
				//    m_ComInterface.SendKeys(keyScanCodes.Length, pKeyReleased, pScanCodes);
				//}
#if false
				bool bArrayKeyUp = true; //Key Released
				int lKeyData = 0x1d; //Scan Codes
				try
				{
					rdpClient.ClientNonScriptable3.SendKeys(1, ref bArrayKeyUp, ref lKeyData);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
#endif
			}
		}

		void StartKeepTimer(Server rdcServer)
		{
			//https://stackoverflow.com/questions/1416803/system-timers-timer-vs-system-threading-timer
			//System.Windows.Forms wraps a native message-only-HWND and uses Window Timers to raise events in that HWNDs message loop.
			//var keepAliveTimer = new KeepAliveTimer();
			var keepAliveTimer = new Timer();
			keepAliveTimer.Interval = timerInterval;
			keepAliveTimer.Tick += new EventHandler(KeepAliveTimer_Tick);
			keepAliveTimer.Tag = rdcServer;
			keepAliveTimer.Start();
			if (serverTimers.TryGetValue(rdcServer, out Timer timer))
			{
				timer.Stop();
			}
			serverTimers[rdcServer] = keepAliveTimer;
		}

		/// <summary>
		/// IMessageFilter allows an application to capture a message before it is dispatched to a control or form.
		/// </summary>
		/// <param name="message">The message to be dispatched. You cannot modify this message.</param>
		/// <returns>
		/// true to filter the message and stop it from being dispatched;
		/// false to allow the message to continue to the next filter or control.
		/// </returns>
		public bool PreFilterMessage(ref Message message)
		{
			Control control = Control.FromHandle(message.HWnd);
			//Debug.WriteLine("{0}|PreFilterMessage|{1},{2},{3}", DateTime.Now.ToString("o"), message.HWnd, message.Msg, control?.Text);
			return false;
		}

		/// https://docs.microsoft.com/en-us/dotnet/api/system.windows.window.isactive?view=netcore-3.1
		/// An active window is the user's current foreground window and has the focus, which is signified by the
		/// active appearance of the title bar. An active window will also be the top-most of all top-level windows
		/// that don't explicitly set the Topmost property.
		bool IsActiveWindow(Control control)
		{
			var hWndForeground = GetForegroundWindow();
			var hWndActive = GetActiveWindow();
			var form = control.FindForm();

			Debug.WriteLine("{0}|{1:x}|IsActiveWindow|Control={2:x}, Foreground={3:x}, Active={4:x}, Form={5:x}", DateTime.Now.ToString("o"), GetCurrentThreadId(), control.Handle.ToInt64(), hWndForeground.ToInt64(), hWndActive.ToInt64(), form.Handle.ToInt64());

			return ((from f in form.MdiChildren select f.Handle)
				.Union(from f in form.OwnedForms select f.Handle)
				.Union(new IntPtr[] { form.Handle })).Contains(hWndForeground);
		}

		private void KeepAliveTimer_Tick(object sender, EventArgs e)
		{
			//string data = ((CustomDerivedTimer)sender).Data;
			Debug.WriteLine("{0}|{1:x}|KeepAliveTimer_Tick", DateTime.Now.ToString("o"), GetCurrentThreadId());

			//https://stackoverflow.com/questions/59624421/rdp-activex-sendkeys-winl-to-lock-screen
			//https://stackoverflow.com/questions/1069990/keep-alive-code-fails-with-new-rdp-client
			var timer = (Timer)sender;
			var rdcServer = (Server)timer.Tag;

			//const uint WM_ACTIVATE = 0x006;
			//const uint WA_ACTIVE = 1;
			//SendMessage(rdcServer.Handle, WM_ACTIVATE, WA_ACTIVE, 0);

			try
			{
				RdpClient rdpClient = Util.GetRdpClient(rdcServer);

				IsActiveWindow(rdpClient.Control);

				rdpClient.Control.WindowTarget.GetType();
				IntPtr hwnd = rdpClient.Control.Handle;
				Debug.WriteLine("{0}|KeepAliveTimer_Tick|rdpClient.Control.Handle={1:x}", DateTime.Now.ToString("o"), hwnd.ToInt64());
				var keyCodes = new Keys[] { Keys.Scroll }; //Keys.ControlKey, Keys.ShiftKey, Keys.Escape
				Util.SendKeys(keyCodes, rdcServer);
			}
			catch (Exception ex)
			{
				timer.Stop();
				Debug.WriteLine("{0}|{1:x}|KeepAliveTimer_Tick|{2}", DateTime.Now.ToString("o"), GetCurrentThreadId(), ex);
			}

			//https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.sendkeys.send?view=netcore-3.1
			//SendKeys.Send("{HOME}"); //{SCROLLLOCK}
		}
#if false
		public void UnknownThreadCalling()
		{
			if (_myDispatcher.CheckAccess())
			{
				// Calling thread is associated with the Dispatcher
			}

			try
			{
				_myDispatcher.VerifyAccess();
				// Calling thread is associated with the Dispatcher
			}
			catch (InvalidOperationException)
			{
				// Thread can't use dispatcher
			}
		}
#endif
	}

	public class KeepAliveTimer : Timer
	{
		Server server;

		protected override void OnTick(EventArgs e)
		{
			//base.OnTick(e);
			Debug.WriteLine("{0}|{1:x}|OnTick", DateTime.Now.ToString("o"), PluginKeepAlive.GetCurrentThreadId());
		}
	}

	//public class Server : ServerBase {
	//	private RdpClient _client;
	//	internal RdpClient Client => _client;
	//}
	public class Util
	{
		public const string SendKeys_Send = "Send";

		public static readonly FieldInfo F_client = typeof(Server).GetField("_client", BindingFlags.NonPublic | BindingFlags.Instance);
		//Type C_SendKeys = Type.GetType("RdcMan.SendKeys", true, true);
		//Type C_SendKeys = Assembly.GetExecutingAssembly().GetType("SendKeys");
		//var innerType = Assembly.GetExecutingAssembly().GetTypes();
		////.Where(t => t.DeclaringType == typeof(Outer))
		////innerType.First(t => t.Name == "SendKeys");
		//foreach(Type t in typeof(IPlugin).Assembly.GetTypes()) Debug.WriteLine(t);
		public static readonly Type C_SendKeys = typeof(IPlugin).Assembly.GetType("RdcMan.SendKeys"); //Assembly.LoadFile(@"RDCMan.exe").GetType("RdcMan.SendKeys");
		public static readonly MethodInfo M_SendKeys = C_SendKeys.GetMethod(SendKeys_Send);

		public static RdpClient GetRdpClient(Server server)
		{
#if false
			//C# 4.0 (released alongside the .NET Framework 4.0 and Visual Studio 2010)
			dynamic rdcServer = obj.Server; // private RdpClient _client;
			//'RDCMan.exe' (CLR v4.0.30319: RDCMan.exe): Loaded 'C:\WINDOWS\Microsoft.Net\assembly\GAC_MSIL\System.Dynamic\v4.0_4.0.0.0__b03f5f7f11d50a3a\System.Dynamic.dll'. Skipped loading symbols. Module is optimized and the debugger option 'Just My Code' is enabled.
			//'RDCMan.exe' (CLR v4.0.30319: RDCMan.exe): Loaded 'Anonymously Hosted DynamicMethods Assembly'. 
			//Exception thrown: 'Microsoft.CSharp.RuntimeBinder.RuntimeBinderException' in System.Core.dll
			//An exception of type 'Microsoft.CSharp.RuntimeBinder.RuntimeBinderException' occurred in System.Core.dll but was not handled in user code
			//'RdcMan.Server._client' is inaccessible due to its protection level
			RdpClient rdpClient = rdcServer._client; // Resolved at runtime, not compile-time
			return rdpClient;
#else
			return F_client.GetValue(server) as RdpClient;
#endif
		}

#if false //internal class SendKeys
		public unsafe static void Send(Keys[] keyCodes, ServerBase serverBase)
		{
			Server serverNode = serverBase.ServerNode;
			RdpClient client = serverNode.Client;
			IMsRdpClientNonScriptable msRdpClientNonScriptable = (IMsRdpClientNonScriptable)client.GetOcx();
			int num = keyCodes.Length;
			try
			{
				SendKeysData sendKeysData = default(SendKeysData);
				bool* ptr = (bool*)sendKeysData.keyUp;
				int* ptr2 = sendKeysData.keyData;
				int num2 = 0;
				for (int i = 0; i < num && i < 10; i++)
				{
					int num3 = (int)Util.MapVirtualKey((uint)keyCodes[i], 0u);
					sendKeysData.keyData[num2] = num3;
					sendKeysData.keyUp[num2++] = 0;
					if (!IsModifier(keyCodes[i]))
					{
						for (int num4 = num2 - 1; num4 >= 0; num4--)
						{
							sendKeysData.keyData[num2] = sendKeysData.keyData[num4];
							sendKeysData.keyUp[num2++] = 1;
						}
						msRdpClientNonScriptable.SendKeys(num2, ref *ptr, ref *ptr2);
						num2 = 0;
					}
				}
			}
			catch
			{
			}
		}
#else
		//MenuHelper.AddSendKeysMenuItems(MainForm#_sessionRemoteActionsMenuItem, ServerTree.Instance.SelectedNode as ServerBase);
		public static void SendKeys(Keys[] keyCodes, ServerBase serverBase)
		{
			object[] args = new object[] { keyCodes, serverBase };
			M_SendKeys.Invoke(null, args);
			//BindingFlags invokeAttr = BindingFlags.Static | BindingFlags.InvokeMethod | BindingFlags.Public;
			//C_SendKeys.InvokeMember(SendKeys_Send, invokeAttr, null, null, args);
		}
#endif
	}
}
