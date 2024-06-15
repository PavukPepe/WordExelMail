using Microsoft.WindowsAPICodePack.Dialogs;
using Spire.Doc;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace WordExelMail
{
    public partial class word_page : Page
    {
        private string currentFilePath = null;

        public word_page(string? path)
        {
            InitializeComponent();
            if (path != null)
            {
                currentFilePath = path;
                LoadRtfFile(path);
            }
        }

        void SaveRtfFile(string _fileName)
        {
            TextRange range = new TextRange(RTB.Document.ContentStart, RTB.Document.ContentEnd);
            FileStream fSteam = new FileStream("врем.rtf", FileMode.Create);
            range.Save(fSteam, DataFormats.Rtf);
            fSteam.Close();

            Document doc = new Document();
            doc.LoadFromFile("врем.rtf");
            doc.SaveToFile(_fileName);
        }

        void LoadRtfFile(string _fileName)
        {
            Document doc = new Document();
            doc.LoadFromFile(_fileName);
            doc.SaveToFile("врем.rtf");

            TextRange range = new TextRange(RTB.Document.ContentStart, RTB.Document.ContentEnd);
            using (FileStream fStream = new FileStream("врем.rtf", FileMode.Open))
            {
                range.Load(fStream, DataFormats.Rtf);
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            CommonSaveFileDialog saveFileDialog = new CommonSaveFileDialog();
            saveFileDialog.DefaultFileName = "document.docx";
            saveFileDialog.Filters.Add(new CommonFileDialogFilter("docx Files", "*.docx"));

            if (saveFileDialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                SaveRtfFile(saveFileDialog.FileName);
                currentFilePath = saveFileDialog.FileName;
            }
        }

    }
}