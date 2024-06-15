using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static Microsoft.WindowsAPICodePack.Shell.PropertySystem.SystemProperties.System;

namespace WordExelMail
{
    public partial class ExcelPage : Page
    {
        DataTable table;
        public ExcelPage(string? path)
        { 
            InitializeComponent();
            DataContext = this;
            Workbook workbook = new Workbook();
            
            if (path != null)
            {
                workbook.LoadFromFile(path);
                Worksheet sheet = workbook.Worksheets[0];
                CellRange locatedRange = sheet.AllocatedRange;
                table = sheet.ExportDataTable(locatedRange, true);
                main_data.ItemsSource = table.DefaultView;
            }
        }

        public ExcelPage()
        {
            InitializeComponent();
            Workbook wb = new Workbook();
            wb.Worksheets.Clear();
            Worksheet sh = wb.Worksheets.Add("Лист 1");
            CellRange locatedRange = sh.AllocatedRange;
            table = sh.ExportDataTable(locatedRange, true);
            main_data.ItemsSource = new DataView();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var table = main_data.ItemsSource as DataView;
            Workbook wb = new Workbook();
            wb.Worksheets.Clear();
            Worksheet sh = wb.Worksheets.Add("Лист 1");
            sh.InsertDataView(table, true, 1, 1);

            SaveFileDialog dil = new SaveFileDialog();
            if (dil.ShowDialog() == true)
            {
                wb.SaveToFile(dil.FileName + ".xlsx", FileFormat.Version2016);
            }
        }

        private void add_but_Click(object sender, RoutedEventArgs e)
        {
            try { 
            
                table.Columns.Add(new DataColumn(name_place.Text));
                main_data.ItemsSource = table.AsDataView();
            }
            catch { }
        }

    }
}
