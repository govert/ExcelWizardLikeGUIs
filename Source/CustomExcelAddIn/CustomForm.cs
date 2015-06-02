using System.Windows.Forms;
using Application = NetOffice.ExcelApi.Application;

namespace CustomExcelAddIn
{
    public partial class CustomForm : Form
    {
        private readonly ExcelSelectionTracker selectionTracker;

        public static bool overrideWindowStyles = false;

        public CustomForm(Application application)
        {
            InitializeComponent();

            selectionTracker = new ExcelSelectionTracker(application, this, ChangeText);
        }

        private void ChangeText(string address)
        {
            AddressBox.Text = address;
            AddressBox.Focus();
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
