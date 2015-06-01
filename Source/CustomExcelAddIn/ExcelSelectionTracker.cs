using System;
using System.Windows.Forms;
using NetOffice;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using Action = System.Action;
using Application = NetOffice.ExcelApi.Application;

namespace CustomExcelAddIn
{
    public class ExcelSelectionTracker
    {
        private readonly Action<string> callback;
        private readonly Form form;
        private readonly Application application;

        public ExcelSelectionTracker(Application application, Form form, Action<string> callback)
        {
            this.callback = callback;
            this.application = application;
            this.form = form;

            application.SheetSelectionChangeEvent += OnNewSelection;
            form.Closed += Unsubscribe;
        }

        private void OnNewSelection(COMObject sh, Range target)
        {
            try
            {
                form.Invoke(new Action(() => callback(target.Address(false, false, XlReferenceStyle.xlA1, true))));
            }
            catch
            {
            }
        }

        private void Unsubscribe(object sender, EventArgs e)
        {
            try
            {
                application.SheetSelectionChangeEvent -= OnNewSelection;
                form.Closed -= Unsubscribe;
            }
            catch
            {
            }
        }
    }
}
