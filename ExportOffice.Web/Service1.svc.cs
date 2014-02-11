using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ExportOffice.Web
{
    public class Service1 : IService1
    {
        public void DoWork()
        {
        }


        public byte[] DoExportExcel(List<FactModel> facts, List<Columns> headersList)
        {
            string filePath = Environment.CurrentDirectory + "\\" + "Test.xlsx";
            try
            {
                var newFile = new FileInfo(filePath);
                if (newFile.Exists)
                {
                    newFile.Delete();
                }
                var table = CreateDt(facts, headersList);
                CreateExcelFile(newFile, table, facts);

                byte[] buffer;
                using (var stream = new FileStream(filePath, FileMode.Open))
                {
                    buffer = new byte[stream.Length];
                    stream.Read(buffer, 0, (int)stream.Length);
                }

                newFile.Delete();
                return buffer;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void CreateExcelFile(FileInfo newFile, DataTable table, List<FactModel> facts)
        {
            using (var package = new ExcelPackage(newFile))
            {
                var worksheet = package.Workbook.Worksheets.Add("Load From DataTable");
                worksheet.Cells["A1"].LoadFromDataTable(table, true, TableStyles.Medium6);

                //var tbl = worksheet.Tables[0];
                //tbl.ShowTotal = true;
                //tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;

                worksheet.Column(1).Width = 14;
                worksheet.Column(2).Width = 12;

                worksheet.HeaderFooter.OddHeader.CenteredText = "مثالی از نحوه‌ی استفاده از ایی پی پلاس";

                worksheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells["B1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                var worksheet1 = package.Workbook.Worksheets.Add("Load From Collection");
                worksheet1.Cells["A2"].LoadFromCollection<FactModel>(facts);

                worksheet1.Column(1).Width = 14;
                worksheet1.Column(2).Width = 12;

                worksheet1.HeaderFooter.OddHeader.CenteredText = "مثالی از نحوه‌ی استفاده از ایی پی پلاس";

                worksheet1.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet1.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                package.Workbook.Properties.Author = "Z.saffarpour";
                package.Workbook.Properties.Title = "Export Excel";
                package.Workbook.Properties.Subject = "Export Excel";
                package.Workbook.Properties.Category = "Sample";
                package.Workbook.Properties.Company = "Zahra Saffarpour";

                worksheet.View.PageBreakView = true;
                worksheet.View.ZoomScale = 110;
                worksheet.View.PageLayoutView = true;
                worksheet.View.RightToLeft = true;

                worksheet1.View.PageLayoutView = true;
                worksheet1.View.RightToLeft = true;

                package.Save();
            }
        }

        private static DataTable CreateDt(IEnumerable<FactModel> facts, IEnumerable<Columns> headersList)
        {
            var table = new DataTable();

            foreach (var headerList in headersList)
            {
                table.Columns.Add(headerList.Header, headerList.ColumnType.GetType());
            }

            foreach (var factModel in facts)
            {

                table.Rows.Add(factModel.Month, factModel.Cost);
            }
            return table;
        }

        public bool DoUploadFile(byte[] buffer)
        {
            try
            {
                string filePath = Environment.CurrentDirectory + "\\" + "Form.Docx";
                using (var stream = new FileStream(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite))
                {
                    stream.Write(buffer, 0, (int)buffer.Length);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public byte[] DoExportWord()
        {
            string filePath = Environment.CurrentDirectory + "\\" + "Form.Docx";
            try
            {
                //byte[] bufferRead;
                //using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))
                //{
                //    bufferRead = new byte[stream.Length];
                //    stream.Read(bufferRead, 0, (int)bufferRead.Length);
                //}
                //using (MemoryStream stream1 = new MemoryStream(bufferRead, true))
                {
                    //WordprocessingDocument wdoc = WordprocessingDocument.Open(filePath, true);
                    using (WordprocessingDocument wdoc = WordprocessingDocument.Open(filePath, true))
                    {
                        var bookMarks = FindBookmarks(wdoc.MainDocumentPart.Document);
                        
                        foreach (var end in bookMarks)
                        {
                            Text textElement = new Text();
                            if (end.Key == "Name" || end.Key == "Name1")
                                textElement = new Text("Zahra");
                            if (end.Key == "Family" || end.Key == "Family1")
                                textElement = new Text("Saffarpour");
                            var runElement = new Run(textElement);

                            end.Value.InsertAfterSelf(runElement);
                        }
                    }
                    //wdoc.Close();

                    byte[] buffer;
                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite))//stream1)
                    {
                        buffer = new byte[stream.Length];
                        stream.Read(buffer, 0, (int)buffer.Length);
                    }

                    return buffer;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ReplaceBookMark(Document file)
        {
            IDictionary<String, BookmarkStart> bookmarkMap = new Dictionary<String, BookmarkStart>();

            foreach (BookmarkStart bookmarkStart in file.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                bookmarkMap[bookmarkStart.Name] = bookmarkStart;
            }

            foreach (BookmarkStart bookmarkStart in bookmarkMap.Values)
            {
                Run bookmarkText = bookmarkStart.NextSibling<Run>();
                if (bookmarkText != null)
                {
                    bookmarkText.GetFirstChild<Text>().Text = "Zahra";
                }
            }
        }

        private static Dictionary<string, BookmarkEnd> FindBookmarks(OpenXmlElement documentPart,
            Dictionary<string, BookmarkEnd> results = null, Dictionary<string, string> unmatched = null)
        {
            results = results ?? new Dictionary<string, BookmarkEnd>();
            unmatched = unmatched ?? new Dictionary<string, string>();

            foreach (var child in documentPart.Elements())
            {
                if (child is BookmarkStart)
                {
                    var bStart = child as BookmarkStart;
                    unmatched.Add(bStart.Id, bStart.Name);
                }

                if (child is BookmarkEnd)
                {
                    var bEnd = child as BookmarkEnd;
                    foreach (var orphanName in unmatched)
                    {
                        if (bEnd.Id == orphanName.Key)
                            results.Add(orphanName.Value, bEnd);
                    }
                }

                FindBookmarks(child, results, unmatched);
            }

            return results;
        }
    }
}
