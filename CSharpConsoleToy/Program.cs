using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

using IDispatch = System.Object;
using IAccessible = System.Object; //System.Windows.Forms.IAccessible

namespace CSharpConsoleToy
{
	class Program
	{
		static readonly Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");

		/// <summary>
		/// Implemented and used by containers and objects to obtain window handles
		/// and manage context-sensitive help.
		/// </summary>
		/// <remarks>
		/// The IOleWindow interface provides methods that allow an application to obtain
		/// the handle to the various windows that participate in in-place activation,
		/// and also to enter and exit context-sensitive help mode.
		/// </remarks>
		[ComImport]
		[Guid("00000114-0000-0000-C000-000000000046")]
		[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
		public interface IOleWindow
		{
			/// <summary>
			/// Returns the window handle to one of the windows participating in in-place activation
			/// (frame, document, parent, or in-place object window).
			/// </summary>
			/// <param name="phwnd">Pointer to where to return the window handle.</param>
			void GetWindow(out IntPtr phwnd);

			/// <summary>
			/// Determines whether context-sensitive help mode should be entered during an
			/// in-place activation session.
			/// </summary>
			/// <param name="fEnterMode"><c>true</c> if help mode should be entered;
			/// <c>false</c> if it should be exited.</param>
			void ContextSensitiveHelp([In, MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
		}

		/// LWSTDAPI IUnknown_GetWindow(IUnknown* punk, HWND *phwnd);
		/// class Marshal { IntPtr GetIUnknownForObject(object o); object GetUniqueObjectForIUnknown (IntPtr unknown); }
		/// #define LWSTDAPI          EXTERN_C DECLSPEC_IMPORT HRESULT STDAPICALLTYPE 
		/// https://docs.microsoft.com/en-us/windows/win32/api/shlwapi/nf-shlwapi-iunknown_getwindow
		[DllImport("shlwapi.dll")] //[MarshalAs(UnmanagedType.Interface)]
		static extern UInt32 IUnknown_GetWindow([In, MarshalAs(UnmanagedType.IUnknown)] object punk, out IntPtr phwnd);

		[DllImport("user32.dll", SetLastError = true)]
		static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

		public delegate bool EnumChildCallback(IntPtr hwnd, ref IntPtr lParam);

		[DllImport("User32.dll")]
		public static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildCallback lpEnumFunc, ref IntPtr lParam);

		[DllImport("User32.dll")]
		public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

		public static bool EnumChildProc(IntPtr hwndChild, ref IntPtr lParam)
		{
			StringBuilder buf = new StringBuilder(128);
			GetClassName(hwndChild, buf, 128);
			if (buf.ToString() == "_WwG")
			{
				lParam = hwndChild;
				return false;
			}
			return true;
		}

		//https://bettersolutions.com/csharp/office-developer-tools/iaccessible.htm
		//[DllImport("oleacc.dll", SetLastError = true)] static extern TODO SystemAccessibleObject(TODO);
		//HRESULT AccessibleObjectFromWindow(HWND hwnd, DWORD dwId, REFIID riid, void** ppvObject);
		[DllImport("Oleacc.dll")]
		static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, out IDispatch ptr);
		[DllImport("oleacc.dll")]
		static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

		// Retrieves the address of the specified interface for the object associated with the specified window. 
		[System.Runtime.InteropServices.DllImport("oleacc.dll", PreserveSig = false, CharSet = System.Runtime.InteropServices.CharSet.Auto, SetLastError = true)]
		[return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Interface)]
		private static extern object AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid);


		// Retrieves the child ID or IDispatch of each child within an accessible container object. 
		[System.Runtime.InteropServices.DllImport("oleacc.dll", CharSet = CharSet.Auto, SetLastError = true)]
		private static extern int AccessibleChildren(
			IAccessible paccContainer,
			int iChildStart,
			int cChildren,
			[Out()] [MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 4)]
			object[] rgvarChildren,
			ref int pcObtained);

#if false
		private IAccessible[] GetAccessibleChildren(IAccessible accContainer)
		{
			// Get the number of child interfaces that belong to this object. 
			int childNum = 0;
			try
			{
				childNum = accContainer.accChildCount;
			}
			catch (Exception ex)
			{
				childNum = 0;
				System.Diagnostics.Debug.Print(ex.Message);
			}

			// Get the child accessible objects. 
			IAccessible[] accObjects = new IAccessible[childNum];
			int count = 0;
			if (childNum != 0)
			{
				AccessibleChildren(accContainer, 0, childNum, accObjects, ref count);
			}
			return accObjects;
		}
#endif 
		internal enum OBJID : uint
		{
			WINDOW = 0x00000000,
			NATIVEOM = 0xFFFFFFF0,
			SYSMENU = 0xFFFFFFFF,
			TITLEBAR = 0xFFFFFFFE,
			MENU = 0xFFFFFFFD,
			CLIENT = 0xFFFFFFFC,
			VSCROLL = 0xFFFFFFFB,
			HSCROLL = 0xFFFFFFFA,
			SIZEGRIP = 0xFFFFFFF9,
			CARET = 0xFFFFFFF8,
			CURSOR = 0xFFFFFFF7,
			ALERT = 0xFFFFFFF6,
			SOUND = 0xFFFFFFF5,

			//public static implicit operator byte(Digit d) => d.digit;
		}

		//https://docs.microsoft.com/en-us/dotnet/csharp/language-reference/operators/user-defined-conversion-operators
		public class OBJIDEnumExt
		{
			//public static implicit operator OBJID(uint value) { return (OBJID)value; }
			//public static implicit operator uint(OBJID value) { return (uint)value; }
			// public static explicit operator xxx(string s) { // details }
			// public static implicit operator string(xxx x)  { // details }
		}

#if false
		//https://stackoverflow.com/questions/2203968/how-to-access-microsoft-word-existing-instance-using-late-binding/2204101#2204101
		static void TestUseLateBindingToGetExcelInstance()
		{
			// Use the window class name ("OpusApp") to retrieve a handle to Word's main window.
			// Alternatively you can get the window handle via the process id:
			// int hwnd = (int)Process.GetProcessById(wordPid).MainWindowHandle;
			//
			IntPtr hwnd = FindWindow("OpusApp", null);

			if (hwnd != null)
			{
				IntPtr hwndChild = null;

				// Search the accessible child window (it has class name "_WwG") 
				// as described in http://msdn.microsoft.com/en-us/library/dd317978%28VS.85%29.aspx
				//
				EnumChildCallback cb = new EnumChildCallback(EnumChildProc);
				EnumChildWindows(hwnd, cb, ref hwndChild);

				if (hwndChild != null)
				{
					// We call AccessibleObjectFromWindow, passing the constant OBJID_NATIVEOM (defined in winuser.h) 
					// and IID_IDispatch - we want an IDispatch pointer into the native object model.
					Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
					IDispatch ptr;

					int hr = AccessibleObjectFromWindow(hwndChild, (uint)OBJID.NATIVEOM, IID_IDispatch.ToByteArray(), out ptr);

					if (hr >= 0)
					{
						object wordApp = ptr.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, ptr, null);

						object version = wordApp.GetType().InvokeMember("Version", BindingFlags.GetField | BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, wordApp, null);
						Console.WriteLine(string.Format("Word version is: {0}", version));
					}
				}
			}
		}
#endif

        //Captain: Input Capture Window, Class: IHWindowClass
        /*
RDCManWithOnlyOneActiveConnections:
	Window 000104EA "sap-ts03 - Remote Desktop Connection Manager v2.7" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
		Window 00010504 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
			Window 00010506 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
		Window 00010508 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
			Window 00010482 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
				Window 00010484 "" WindowsForms10.SCROLLBAR.app.0.3d90434_r6_ad1
				Window 00020486 "" ATL:00007FF8BD54F2F0
				Window 0001051C "Connected to sap-ts03" WindowsForms10.STATIC.app.0.3d90434_r6_ad1
				Window 0012120A "" ATL:00007FF8BD54F2F0
					Window 000D073E "" UIMainClass
						Window 001604DC "" UIContainerClass
							Window 000611DE "Input Capture Window" IHWindowClass
							Window 00070558 "Output Painter Window" OPContainerClass
								Window 001211E4 "Output Painter Child Window" OPWindowClass
								Window 001307E8 "Output Painter DX Child Window" OPWindowClass
			Window 0001050A "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
			Window 0001050C "" WindowsForms10.SysTreeView32.app.0.3d90434_r6_ad1

RDCManWithTwoActiveConnections:
	Window 000104EA "hz-ts03 - Remote Desktop Connection Manager v2.7" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
		Window 00010504 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
			Window 00010506 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
		Window 00010508 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
			Window 00010482 "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
				Window 00010484 "" WindowsForms10.SCROLLBAR.app.0.3d90434_r6_ad1
				Window 00020486 "" ATL:00007FF8BD54F2F0
				Window 0012120A "" ATL:00007FF8BD54F2F0
					Window 000D073E "" UIMainClass
						Window 001604DC "" UIContainerClass
							Window 000611DE "Input Capture Window" IHWindowClass
							Window 00070558 "Output Painter Window" OPContainerClass
								Window 001211E4 "Output Painter Child Window" OPWindowClass
								Window 001307E8 "Output Painter DX Child Window" OPWindowClass
				Window 00020466 "Connected to hz-ts03" WindowsForms10.STATIC.app.0.3d90434_r6_ad1
				Window 00A70C6C "" ATL:00007FF8BD54F2F0
					Window 00130FA8 "" UIMainClass
						Window 0008019E "" UIContainerClass
							Window 000D101C "Input Capture Window" IHWindowClass
							Window 000A102E "Output Painter Window" OPContainerClass
								Window 00090FE8 "Output Painter Child Window" OPWindowClass
			Window 0001050A "" WindowsForms10.Window.8.app.0.3d90434_r6_ad1
			Window 0001050C "" WindowsForms10.SysTreeView32.app.0.3d90434_r6_ad1
		*/
        public static bool EnumRDPChildProc(IntPtr hwndChild, ref IntPtr lParam)
		{
			object ptr = null;
			Guid guid = type.GUID;
			int hr = AccessibleObjectFromWindow(hwndChild, (uint)OBJID.NATIVEOM, ref guid, ref ptr);
			if (hr >= 0)
			{
				Console.WriteLine("ptr=" + ptr);
				Marshal.ReleaseComObject(ptr);
			}
			else
			{
				Console.WriteLine("hr={0:X}, hwndChild={1:X}", hr, hwndChild.ToInt64());
			}
			return true;
		}

		private static object GetAccessibleObject(IntPtr hwnd)
		{
			if (IntPtr.Zero != hwnd)
			{
				IntPtr hWndChild = new IntPtr();
				EnumChildCallback cb = new EnumChildCallback(EnumRDPChildProc);
				EnumChildWindows(hwnd, cb, ref hWndChild);
			}
			return null;
		}

		//https://stackoverflow.com/questions/5549827/how-do-i-get-a-com-interface-given-a-hwnd-of-an-activex-control
		static void Main(string[] args)
		{
			Guid guid = type.GUID; // new Guid("{618736E0-3C3D-11CF-810C-00AA00389B71}");
			//Console.WriteLine(guid); //-2147467262
			//object obj = null;
			//int retVal = AccessibleObjectFromWindow(new IntPtr(0x000611DE), (uint)OBJID.WINDOW, ref guid, ref obj);
			////accessible = (IAccessible)obj;
			//Console.WriteLine(retVal); //-2147467262
			//Console.WriteLine(null != obj ? "good" : "null");
			//throw new COMException("AccessibleObjectFromWindow", retVal);

			GetAccessibleObject(new IntPtr(0x000104EA));
			Console.ReadKey();
		}
		static Type type = Type.GetTypeFromProgID("MsTscAx.MsTscAx");
	}
}
