using CefSharp;
using CefSharp.WinForms;
using ExcelForms3.Consult;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VsNovell
{
    internal class MainWeb
    {
        private static ChromiumWebBrowser _instanceBrowser = null;
        private static Main _instanceMainForm = null;


        public MainWeb(ChromiumWebBrowser originalBrowser, Main mainForm)
        {
            _instanceBrowser = originalBrowser;
            _instanceMainForm = mainForm;
        }

        public void ShowDevTools()
        {
            _instanceBrowser.ShowDevTools();
        }

        public string OpenReadForm(string table)
        {
            ExcelData data = null;
            if(table == "1")
            {
                data = _instanceMainForm.ReadExcelTable1();
            }
            else
            {
                data = _instanceMainForm.ReadExcelTable2();
            }
            return JsonConvert.SerializeObject(data);
        }

        public string EqualsTable()
        {
            return JsonConvert.SerializeObject(_instanceMainForm.EqualsTables());
        }
    }
}
