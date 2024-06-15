using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WordExelMail
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void open_excel_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog openFileDialog = new CommonOpenFileDialog() ;
            openFileDialog.Filters.Add(new CommonFileDialogFilter("", "xlsx"));
            CommonFileDialogResult res = openFileDialog.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                work_space.Content = new ExcelPage(openFileDialog.FileName);
            }

        }

        private void create_excel_Click(object sender, RoutedEventArgs e)
        {

        }

        private void open_word_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog openFileDialog = new CommonOpenFileDialog();
            CommonFileDialogResult res = openFileDialog.ShowDialog();
            if (res == CommonFileDialogResult.Ok)
            {
                work_space.Content = new word_page(openFileDialog.FileName);
            }
        }

        private void create_word_Click(object sender, RoutedEventArgs e)
        {
            work_space.Content = new word_page(null);
        }
    }
}