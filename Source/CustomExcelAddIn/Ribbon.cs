using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Application = NetOffice.ExcelApi.Application;

namespace CustomExcelAddIn
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        private readonly Application application;

        public MyRibbon()
        {
            application = new Application(null, ExcelDnaUtil.Application);
        }

        public void OnButtonPressed1(IRibbonControl control)
        {
            CustomForm form = new CustomForm(application);
            form.Show();
        }

        public void OnButtonPressed2(IRibbonControl control)
        {
            CustomForm form = new CustomForm(application);
            form.ShowDialog();
        }

        public void OnButtonPressed3(IRibbonControl control)
        {
            Thread thread = new Thread(() =>
            {
                CustomForm form = new CustomForm(application);
                form.Show();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        [DllImport("kernel32.dll")]
        public static extern int GetCurrentThreadId();

        public void OnButtonPressed4(IRibbonControl control)
        {
            int id = GetCurrentThreadId();

            Thread thread = new Thread(() =>
            {
                CustomForm form = new CustomForm(application, id);
                form.ShowDialog();
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public void OnButtonPressed5(IRibbonControl control)
        {
            CustomForm form = new CustomForm(application);
            form.Show(new Win32Window(new IntPtr(application.Hwnd)));
        }

        public void OnButtonPressed6(IRibbonControl control)
        {
            CustomForm form = new CustomForm(application);
            form.ShowDialog(new Win32Window(new IntPtr(application.Hwnd)));
        }

        public void OnButtonPressed7(IRibbonControl control)
        {
            Thread thread = new Thread(() =>
            {
                CustomForm form = new CustomForm(application);
                form.Show(new Win32Window(new IntPtr(application.Hwnd)));
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public void OnButtonPressed8(IRibbonControl control)
        {
            Thread thread = new Thread(() =>
            {
                CustomForm form = new CustomForm(application);
                form.ShowDialog(new Win32Window(new IntPtr(application.Hwnd)));
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        private delegate bool EnumWindowsCallback(IntPtr hwnd, IntPtr param);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsCallback callback, IntPtr param);

        [DllImport("user32.dll")]
        private static extern int GetClassNameW(IntPtr hwnd, [MarshalAs(UnmanagedType.LPWStr)] StringBuilder buf, int nMaxCount);

        private static IntPtr FindWorkbookWindow(IntPtr mainWindowHandle)
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

        public void OnButtonPressed9(IRibbonControl control)
        {
            CustomForm form = new CustomForm(application);

            IntPtr hWndChild = FindWorkbookWindow(new IntPtr(application.Hwnd));

            form.Show(new Win32Window(hWndChild));
        }

        public void OnCheckBoxPressed1(IRibbonControl control, bool pressed)
        {
            CustomForm.overrideWindowStyles = pressed;
        }

        public override string GetCustomUI(string uiName)
        {
            return File.ReadAllText("ribbon.xml");
        }
    }
}
