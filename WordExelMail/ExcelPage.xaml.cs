using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Xls;
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

namespace WordExelMail
{
    public partial class ExcelPage : Page
    {
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
                var table = sheet.ExportDataTable(locatedRange, true);
                main_data.ItemsSource = table.DefaultView;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var table = main_data.ItemsSource as DataView;
            Workbook wb = new Workbook();
            wb.Worksheets.Clear();
            Worksheet sh = wb.Worksheets.Add("Лист 1");
            sh.InsertDataView(table, true, 1, 1);
            CommonSaveFileDialog dil = new CommonSaveFileDialog();
            
            if (dil.ShowDialog() == CommonFileDialogResult.Ok)
            {
                MessageBox.Show(dil.FileName);
                wb.SaveToFile(dil.FileName, FileFormat.Version2016);
            }
        }
    }
}
