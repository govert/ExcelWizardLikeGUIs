using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Application = NetOffice.ExcelApi.Application;

namespace CustomExcelAddIn
{
    public partial class CustomForm : Form
    {
        public delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        private int hHook = 0;
        private const int WH_CALLWNDPROC = 4;
        private const int WH_MOUSE = 7;
        private const int WH_GETMESSAGE = 3;
        private const int WH_MOUSE_LL = 14;
        private const int WM_MOUSEACTIVATE = 0x0021;
        private const int MA_ACTIVATE = 1;
        private const int MA_ACTIVATEANDEAT = 2;
        private const int MA_NOACTIVATE = 3;
        private const int MA_NOACTIVATEANDEAT = 4;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONUP = 0x0202;
        private const int WM_MOUSEMOVE = 0x0200;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_RBUTTONUP = 0x0205;
        private const int WM_ACTIVATE = 6;
        
        [StructLayout(LayoutKind.Sequential)]
        public class POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public class MouseHookStruct
        {
            public POINT pt;
            public int hwnd;
            public int wHitTestCode;
            public int dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CWPSTRUCT
        {
            public IntPtr lparam;
            public IntPtr wparam;
            public int message;
            public IntPtr hwnd;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        public static extern int GetCurrentThreadId();

        [DllImport("kernel32.dll")]
        static extern uint GetLastError();

        private readonly ExcelSelectionTracker selectionTracker;

        public static bool overrideWindowStyles = false;
        private readonly Application app;

        private delegate bool EnumWindowsCallback(IntPtr hwnd, IntPtr param);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsCallback callback, IntPtr param);

        [DllImport("user32.dll")]
        private static extern int GetClassNameW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);

        private readonly IntPtr workbookHandle;

        public static IntPtr FindWorkbookWindow(IntPtr mainWindowHandle)
        {
            IntPtr hWndChild = IntPtr.Zero;

            StringBuilder cname = new StringBuilder(256);
            EnumChildWindows(mainWindowHandle, delegate(IntPtr hWndEnum, IntPtr param)
            {
                GetClassNameW(hWndEnum, cname, cname.Capacity);
                if (cname.ToString() == "EXCEL7")
                {
                    hWndChild = hWndEnum;
                    return false;
                }
                return true;
            }, IntPtr.Zero);

            return hWndChild;
        }

        public CustomForm(Application application, int xlThreadId = 0)
        {
            InitializeComponent();
            app = application;

            selectionTracker = new ExcelSelectionTracker(application, this, ChangeText);

            workbookHandle = FindWorkbookWindow(new IntPtr(app.Hwnd));


            if (xlThreadId != 0)
            {
                HookProc proc = MouseHookProc;
                //hHook = SetWindowsHookEx(WH_MOUSE_LL, proc, (IntPtr) 0, 0);
                hHook = SetWindowsHookEx(WH_MOUSE, proc, (IntPtr)0, xlThreadId);

                //HookProc proc = MsgProc;
                //hHook = SetWindowsHookEx(WH_GETMESSAGE, proc, (IntPtr)0, xlThreadId);

                //HookProc proc = CwpProc;
                //hHook = SetWindowsHookEx(WH_CALLWNDPROC, proc, (IntPtr)0, xlThreadId);
            }
        }

        //protected override void WndProc(ref Message m)
        //{
        //    if (m.Msg == WM_MOUSEACTIVATE)
        //    {
        //        Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE  WndProc - res=" + m.Result);
        //        //m.Result = new IntPtr(4);
        //        //return;
        //    }

        //    base.WndProc(ref m);

        //    if (m.Msg == WM_MOUSEACTIVATE)
        //    {
        //        Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE  WndProc - res=" + m.Result);
        //        //m.Result = new IntPtr(4);
        //        //return;
        //    }
        //}

        public int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            MouseHookStruct mouseHookStruct = (MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseHookStruct));

            if (nCode < 0)
            {
                return CallNextHookEx(hHook, nCode, wParam, lParam);
            }

            if (wParam.ToInt32() == WM_MOUSEACTIVATE)
            {
                Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE");
                int ret = CallNextHookEx(hHook, nCode, wParam, lParam);
                Debug.Print(DateTime.Now.Ticks + " CallNextHookEx = " + ret);
                return ret;
            }

            return CallNextHookEx(hHook, nCode, wParam, lParam);
        }

        public int CwpProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            CWPSTRUCT cwpStruct = (CWPSTRUCT)Marshal.PtrToStructure(lParam, typeof(CWPSTRUCT));

            if (nCode < 0)
            {
                return CallNextHookEx(hHook, nCode, wParam, lParam);
            }

            if (cwpStruct.hwnd == workbookHandle)
            {
                Debug.Print(DateTime.Now.Ticks + " cwpStruct.message=" + cwpStruct.message);
            }

            if (wParam.ToInt32() == WM_MOUSEACTIVATE)
            {
                Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE");
            }

            if (cwpStruct.message == WM_MOUSEACTIVATE)
            {
                Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE");
            }
          
            return CallNextHookEx(hHook, nCode, wParam, lParam);
        }

        public int MsgProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            Message msg = (Message)Marshal.PtrToStructure(lParam, typeof(Message));

            if (nCode < 0)
            {
                return CallNextHookEx(hHook, nCode, wParam, lParam);
            }

            if (msg.Msg == WM_LBUTTONDOWN)
            {
                Debug.Print(DateTime.Now.Ticks + " msg.Msg=" + msg.Msg);
            }

            return CallNextHookEx(hHook, nCode, wParam, lParam);
        }

        private void ChangeText(string address)
        {
            AddressBox.Text = address;
            AddressBox.Select(address.Length, 0);
            BringToFront();
            Activate();
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams ret = base.CreateParams;

                if (overrideWindowStyles)
                {
                    ret.ExStyle |= (int)WindowStyles.WS_EX_NOACTIVATE;
                }

                return ret;
            }
        }
    }
}
