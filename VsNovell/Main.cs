using CefSharp;
using CefSharp.WinForms;
using ExcelForms3.Consult;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VsNovell
{
    public partial class Main : Form
    {
        public ChromiumWebBrowser MainWeb { get; set; }
        public Main()
        {
            InitializeComponent();
            InitializeChromium();

            MainWeb meinWeb = new MainWeb(MainWeb, this);


            CefSharpSettings.WcfEnabled = true;
            MainWeb.JavascriptObjectRepository.Settings.LegacyBindingEnabled = true;
            MainWeb.JavascriptObjectRepository.Register("mainWeb", meinWeb, isAsync: false, options: BindingOptions.DefaultBinder);
        }

        public void InitializeChromium()
        {
            CefSettings settings = new CefSettings();
            Cef.Initialize(settings);

            string path = string.Format("{0}/MainWeb/Index/index.html", Environment.CurrentDirectory);

            MainWeb = new ChromiumWebBrowser(path);
            this.Controls.Add(MainWeb);
            MainWeb.Dock = DockStyle.Fill;
        }




        ExcelData table1;
        ExcelData table2;

        public string OpenRead()
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Excel file (*.xlsx)|*.xlsx|Old excel file (*.xls)|*.xls";

            if (open.ShowDialog() == DialogResult.OK)
            {
                return open.FileName;
            }
            return null;
        }

        public ExcelData ReadExcelTable1()
        {
            ExcelData result = null;

            this.Invoke(new Action(() => {

                string path = OpenRead();
                if (path != null)
                {
                    ExcelRead excel = new ExcelRead(path);
                    excel.ReadDokument(new ExcelCell(5), new ExcelCell(6), new ExcelCell(9));
                    table1 = result = excel.ReadData;
                }
            }));
            return result;
        }
        public ExcelData ReadExcelTable2()
        {
            ExcelData result = null;

            this.Invoke(new Action(() => {

                string path = OpenRead();
                if (path != null)
                {
                    ExcelRead excel = new ExcelRead(path);
                    excel.ReadDokument(new ExcelCell(4), new ExcelCell(6), new ExcelCell(10, "<p_menge>", "</p_menge>"));
                    table2 = result = excel.ReadData;
                }
            }));
            return result;
        }






        public EqualsExcelTable EqualsTables()
        {
            EqualsExcelTable result = new EqualsExcelTable();
            //richTextBox3.Text += "Ошибочно вбиты в таблицу 1:\n";
            result.ErrorTable1 = DetectNotEuqals(table1, table2);
            //richTextBox3.Text += "Ошибочно вбиты в таблицу 2:\n";
            result.ErrorTable2 = DetectNotEuqals(table2, table1);

            //richTextBox3.Text += "Полные совпадения :\n";
            result.Contains = DetectEuqals(table2, table1);

            result.RestTable1 = Convert(table1);
            result.RestTable2 = Convert(table2);

            return result;
            //richTextBox3.Text += "Осталось 1:\n";
            //ShowTable(table1);

            //richTextBox3.Text += "Осталось 2:\n";
            //ShowTable(table2);
        }

        public List<string[]> Convert(ExcelData table)
        {
            List<string[]> bacet = new List<string[]>();

            foreach (var stroka in table1.Data)
            {
                bacet.Add(stroka);
            }
            return bacet;
        }

        public List<string[]> DetectNotEuqals(ExcelData table1, ExcelData table2)
        {
            List<string[]> bacet = new List<string[]>();

            foreach (var stroka in table1.Data)
            {
                var find = table2.Data.FirstOrDefault(str => str[0] + str[1] == stroka[0] + stroka[1]);
                if (find == null)
                {
                    bacet.Add(stroka);
                }
            }

            foreach (var item in bacet)
            {
                table1.Data.Remove(item);
            }

            return bacet;
        }

        public List<string[]> DetectEuqals(ExcelData table1, ExcelData table2)
        {
            List<string[]> bacet = new List<string[]>();

            foreach (var stroka in table1.Data)
            {
                var find = table2.Data.FirstOrDefault(str => str[0] + str[1] + str[2] == stroka[0] + stroka[1] + stroka[2]);
                if (find != null)
                {
                    bacet.Add(stroka);
                    table2.Data.Remove(find);
                }
            }

            foreach (var item in bacet)
            {
                table1.Data.Remove(item);
            }

            return bacet;
        }

    }

    public class EqualsExcelTable
    {
        public List<string[]> ErrorTable1 { get; set; }
        public List<string[]> ErrorTable2 { get; set; }
        public List<string[]> Contains { get; set; }
        public List<string[]> RestTable1 { get; set; }
        public List<string[]> RestTable2 { get; set; }

    }
}
