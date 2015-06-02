using System;
using System.IO;
using System.Runtime.InteropServices;
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

        public void OnButtonPressed4(IRibbonControl control)
        {
            Thread thread = new Thread(() =>
            {
                CustomForm form = new CustomForm(application);
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
