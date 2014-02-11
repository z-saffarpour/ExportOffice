using ExportOffice.ServiceReference1;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;

namespace ExportOffice.ViewModel
{
    public delegate void ExportExcelCompleted(byte[] buffer);
    public delegate void ExportWordCompleted(byte[] buffer);
    public delegate void UploadCompleted();
    public class FactViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<FactModel> FactModels = new ObservableCollection<FactModel>();
        private ObservableCollection<Columns> Columnss = new ObservableCollection<Columns>();
        public ExportExcelCompleted ExportExcelCompleted;
        public ExportWordCompleted ExportWordCompleted;
        public UploadCompleted UploadCompleted;
        public FactViewModel()
        {
            CreateDt();
        }
        private void CreateDt()
        {
            Columnss.Add(new Columns { Header = "ماه", ColumnType = "string" });
            Columnss.Add(new Columns { Header = "هزینه", ColumnType = "decimal" });

            FactModels.Add(new FactModel { Month = "فروردین", Cost = 100 });
            FactModels.Add(new FactModel { Month = "اردیبهشت", Cost = 250 });
            FactModels.Add(new FactModel { Month = "خرداد", Cost = 80 });
            FactModels.Add(new FactModel { Month = "تیر", Cost = 300 });
            FactModels.Add(new FactModel { Month = "مرداد", Cost = 200 });
            FactModels.Add(new FactModel { Month = "شهریور", Cost = 150 });
            FactModels.Add(new FactModel { Month = "مهر", Cost = 250 });
            FactModels.Add(new FactModel { Month = "آبان", Cost = 200 });
            FactModels.Add(new FactModel { Month = "آذر", Cost = 400 });
            FactModels.Add(new FactModel { Month = "دی", Cost = 100 });
            FactModels.Add(new FactModel { Month = "بهمن", Cost = 130 });
            FactModels.Add(new FactModel { Month = "اسفند", Cost = 80 });
        }

        public void CreateExcelFile()
        {
            var srv = new Service1Client();
            srv.DoExportExcelAsync(FactModels, Columnss);
            srv.DoExportExcelCompleted += srv_DoExportExcelCompleted;
            srv.CloseAsync();
        }

        void srv_DoExportExcelCompleted(object sender, DoExportExcelCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                var buffer = e.Result;
                if (ExportExcelCompleted != null)
                    ExportExcelCompleted(buffer);
            }
            else
                MessageBox.Show(e.Error.Message);
        }

        public void UploadForm(byte[] buffer)
        {
            var srv = new Service1Client();
            srv.DoUploadFileAsync(buffer);
            srv.DoUploadFileCompleted += srv_DoUploadFileCompleted;
            srv.CloseAsync();
        }

        void srv_DoUploadFileCompleted(object sender, DoUploadFileCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                MessageBox.Show("آپلود فایل با موفقیت انجام شد.", "اعلام", MessageBoxButton.OK);
                if (UploadCompleted != null)
                    UploadCompleted();
            }
            else
                MessageBox.Show(e.Error.Message);
        }

        public void ExportFormWord()
        {
            var srv = new Service1Client();
            srv.DoExportWordAsync();
            srv.DoExportWordCompleted += srv_DoExportWordCompleted;
            srv.CloseAsync();
        }

        void srv_DoExportWordCompleted(object sender, DoExportWordCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                var buffer = e.Result;
                if (ExportWordCompleted != null)
                    ExportWordCompleted(buffer);
            }
            else
                MessageBox.Show(e.Error.Message);
        }
    }
}
