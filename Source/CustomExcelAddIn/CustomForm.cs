using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = NetOffice.ExcelApi.Application;

namespace CustomExcelAddIn
{
    public partial class CustomForm : Form
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private readonly ExcelSelectionTracker selectionTracker;

        public static bool overrideWindowStyles = false;
        private readonly Application app;

        public CustomForm(Application application, bool trackMouse = false)
        {
            InitializeComponent();
            app = application;

            selectionTracker = new ExcelSelectionTracker(application, this, ChangeText);

            if (trackMouse)
            {
                MouseLeave += CheckMouse;
                MouseEnter += CheckMouse;
            }
        }


        private void CheckMouse(object sender, EventArgs e)
        {
            if (ClientRectangle.Contains(PointToClient(MousePosition)))
            {
                BringToFront();
                Activate();
            }
            else
            {
                try
                {
                    SetForegroundWindow(new IntPtr(app.Hwnd));
                }
                catch
                {
                }
            }
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
                    ret.ExStyle |= (int) WindowStyles.WS_EX_NOACTIVATE;
                }

                return ret;
            }
        }
    }
}
