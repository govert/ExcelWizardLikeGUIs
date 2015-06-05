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
        //public delegate int CbtProc(CbtEvents nCode, IntPtr wParam, IntPtr lParam);

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

        /// <summary>
        /// These are the different hook types when registering a callback with SetWindowsHookEx().
        /// For more details, see http://msdn.microsoft.com/en-us/library/ms644990. </summary>
        public enum HookType
        {
            WH_JOURNALRECORD = 0,
            WH_JOURNALPLAYBACK = 1,
            WH_KEYBOARD = 2,
            WH_GETMESSAGE = 3,
            WH_CALLWNDPROC = 4,
            WH_CBT = 5,
            WH_SYSMSGFILTER = 6,
            WH_MOUSE = 7,
            WH_HARDWARE = 8,
            WH_DEBUG = 9,
            WH_SHELL = 10,
            WH_FOREGROUNDIDLE = 11,
            WH_CALLWNDPROCRET = 12,
            WH_KEYBOARD_LL = 13,
            WH_MOUSE_LL = 14
        }


        /// <summary>
        /// These are the message IDs that can be passed in as the code parameter
        /// of the WindowsHookCallback that was registered with HookType.WH_CBT.
        /// For details, see http://msdn.microsoft.com/en-us/library/ms644977. </summary>
        public enum CbtEvents
        {
            HCBT_MOVESIZE = 0,
            HCBT_MINMAX = 1,
            HCBT_QS = 2,
            HCBT_CREATEWND = 3,
            HCBT_DESTROYWND = 4,
            HCBT_ACTIVATE = 5,
            HCBT_CLICKSKIPPED = 6,
            HCBT_KEYSKIPPED = 7,
            HCBT_SYSCOMMAND = 8,
            HCBT_SETFOCUS = 9
        }

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

        [StructLayout(LayoutKind.Sequential)]
        public struct CWPRETSTRUCT
        {
            public int lresult;
            public IntPtr lparam;
            public IntPtr wparam;
            public int message;
            public IntPtr hwnd;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(HookType idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern IntPtr SetFocus(IntPtr hWnd);

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

        HookProc _procMsg;
        HookProc _procCwp;
        HookProc _procCwpRet;
        HookProc _procCbt;

        private int hHookMsg = 0;
        private int hHookCwp = 0;
        private int hHookCwpRet = 0;
        private int hHookCbt = 0;


        public CustomForm(Application application, int xlThreadId = 0)
        {
            InitializeComponent();
            app = application;

            selectionTracker = new ExcelSelectionTracker(application, this, ChangeText);

            workbookHandle = FindWorkbookWindow(new IntPtr(app.Hwnd));


            if (xlThreadId != 0)
            {
                _procMsg = MsgProc;
                _procCwp = CwpProc;
                _procCwpRet = CallWndRetProc;
                _procCbt = CbtProc;

                //HookProc proc = MouseHookProc;
                //hHook = SetWindowsHookEx(WH_MOUSE_LL, proc, (IntPtr) 0, 0);
                //hHook = SetWindowsHookEx(WH_MOUSE, proc, (IntPtr)0, xlThreadId);

                hHookMsg = SetWindowsHookEx(HookType.WH_GETMESSAGE, _procMsg, (IntPtr)0, xlThreadId);
                hHookCwp = SetWindowsHookEx(HookType.WH_CALLWNDPROC, _procCwp, (IntPtr)0, xlThreadId);
                hHookCwpRet = SetWindowsHookEx(HookType.WH_CALLWNDPROCRET, _procCwpRet, (IntPtr)0, xlThreadId);
                hHookCbt = SetWindowsHookEx(HookType.WH_CBT, _procCwpRet, (IntPtr)0, xlThreadId);
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

        //public int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        //{
        //    MouseHookStruct mouseHookStruct = (MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(MouseHookStruct));

        //    if (nCode < 0)
        //    {
        //        return CallNextHookEx(hHook, nCode, wParam, lParam);
        //    }

        //    if (wParam.ToInt32() == WM_MOUSEACTIVATE)
        //    {
        //        Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE");
        //        int ret = CallNextHookEx(hHook, nCode, wParam, lParam);
        //        Debug.Print(DateTime.Now.Ticks + " CallNextHookEx = " + ret);
        //        return ret;
        //    }

        //    return CallNextHookEx(hHook, nCode, wParam, lParam);
        //}

        public int CwpProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            CWPSTRUCT cwpStruct = (CWPSTRUCT)Marshal.PtrToStructure(lParam, typeof(CWPSTRUCT));

            if (nCode < 0)
            {
                return CallNextHookEx(hHookCwp, nCode, wParam, lParam);
            }

            Debug.Assert(nCode == 0 /* HC_ACTION */);

            if (cwpStruct.hwnd == workbookHandle)
            {
                // Debug.Print(DateTime.Now.Ticks + " cwpStruct.message=" + cwpStruct.message);
            }

            if (cwpStruct.message == WM_MOUSEACTIVATE)
            {
                Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE");
                SetFocus(cwpStruct.hwnd);
            }
          
            return CallNextHookEx(hHookCwp, nCode, wParam, lParam);
        }

        public int CallWndRetProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0 || lParam == IntPtr.Zero)
            {
                return CallNextHookEx(hHookCwpRet, nCode, wParam, lParam);
            }

            CWPRETSTRUCT cwpRetStruct = (CWPRETSTRUCT)Marshal.PtrToStructure(lParam, typeof(CWPRETSTRUCT));
            if (cwpRetStruct.hwnd == workbookHandle)
            {
                // Debug.Print(DateTime.Now.Ticks + " cwpStruct.message=" + cwpRetStruct.message);
            }

            if (cwpRetStruct.message == WM_MOUSEACTIVATE)
            {
                Debug.Print(DateTime.Now.Ticks + "WM_MOUSEACTIVATE RETURN: " + cwpRetStruct.lresult);
                cwpRetStruct.lresult = 1;
                Marshal.StructureToPtr(cwpRetStruct, lParam, true);
            }

            return CallNextHookEx(hHookCwpRet, nCode, wParam, lParam);
        }

        public int MsgProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            Message msg = (Message)Marshal.PtrToStructure(lParam, typeof(Message));
            //Debug.Print(msg.ToString());

            if (msg.Msg == WM_MOUSEACTIVATE)
            {
                Debug.Print(DateTime.Now.Ticks + " WM_MOUSEACTIVATE");
            }

            if (nCode < 0)
            {
                return CallNextHookEx(hHookMsg, nCode, wParam, lParam);
            }

            if (msg.Msg == WM_LBUTTONDOWN)
            {
                Debug.Print(DateTime.Now.Ticks + " msg.Msg=" + msg.Msg);
            }

            return CallNextHookEx(hHookMsg, nCode, wParam, lParam);
        }

        public int CbtProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //Message msg = (Message)Marshal.PtrToStructure(lParam, typeof(Message));
            //Debug.Print(msg.ToString());

            CbtEvents code = (CbtEvents)nCode;
            if (code == CbtEvents.HCBT_CLICKSKIPPED)
            {
                Debug.Print(">>>>>>>>>>>>>");
            }

            return CallNextHookEx(hHookCbt, nCode, wParam, lParam);
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
