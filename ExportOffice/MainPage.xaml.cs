using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Linq;
namespace ExportOffice
{
    public partial class MainPage : UserControl
    {
        readonly ViewModel.FactViewModel _vm = new ViewModel.FactViewModel();
        public MainPage()
        {
            InitializeComponent();
        }

        private void Button1_OnClick(object sender, RoutedEventArgs e)
        {
            _vm.ExportExcelCompleted = buffer =>
            {
                var path = Environment.CurrentDirectory + "\\" + "Test.xlsx";
                var newFile = new FileInfo(path);
                if (newFile.Exists)
                    newFile.Delete();

                using (var stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    stream.Write(buffer, 0, (int)buffer.Length);
                }
                MessageBox.Show(string.Format("فایل در مسیر {0} با موفقیت ذخیره شد.", path), "اعلام", MessageBoxButton.OK);
            };
            _vm.CreateExcelFile();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var openFile = new OpenFileDialog();
            //openFile.Filter = "Word File (.Docx)|.docx";
            var dialogResualt = openFile.ShowDialog();
            if (dialogResualt == true)
            {
                var file = openFile.Files.FirstOrDefault();
                if (file != null)
                {
                    byte[] buffer;
                    using (var stream = file.OpenRead())
                    {
                        buffer = new byte[stream.Length];
                        stream.Read(buffer, 0, (int)stream.Length);
                    }

                    if (buffer != null && buffer.Length > 0)
                    {
                        _vm.UploadCompleted = () =>
                        {
                            ExportWord.IsEnabled = true;
                        };
                        _vm.UploadForm(buffer);
                    }
                }
            }
            //_vm.UploadForm()
        }

        private void ExportWord_Click(object sender, RoutedEventArgs e)
        {
            _vm.ExportWordCompleted = buffer =>
            {
                var path = Environment.CurrentDirectory + "\\" + "Test.Docx";
                var newFile = new FileInfo(path);
                if (newFile.Exists)
                    newFile.Delete();

                using (var stream = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite))
                {
                    stream.Write(buffer, 0, (int)buffer.Length);
                }
                MessageBox.Show(string.Format("فایل در مسیر {0} با موفقیت ذخیره شد.", path), "اعلام", MessageBoxButton.OK);
            };
            _vm.ExportFormWord();
        }
    }
}
