using System;
using System.Windows.Forms;
using Application = NetOffice.ExcelApi.Application;

namespace CustomExcelAddIn
{
    public partial class CustomForm : Form
    {
        private readonly ExcelSelectionTracker selectionTracker;

        public CustomForm(Application application)
        {
            InitializeComponent();

            selectionTracker = new ExcelSelectionTracker(application, this, ChangeText);
        }

        private void ChangeText(string address)
        {
            AddressBox.Text = address;
        }
    }
}
