using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.IO;

namespace DioDocsBenchmarkApp1
{
    public class MyConfig : ManualConfig
    {
        public MyConfig()
        {
            WithOptions(ConfigOptions.DisableOptimizationsValidator);
        }
    }

    [Config(typeof(MyConfig))]
    [SimpleJob(RuntimeMoniker.Net50)]
    public class Benchmark
    {
        [GlobalSetup]
        public static void Setup() => GrapeCity.Documents.Excel.Workbook.SetLicenseKey("");

        private const int ColumnNum = 100;
        private const int RowNum = 100;

        [Benchmark]
        public void DioDocs()
        {
            var workbook = new GrapeCity.Documents.Excel.Workbook();
            var worksheet = workbook.ActiveSheet;
            for (var i = 1; i <= ColumnNum; i++)
            {
                for (var j = 1; j <= RowNum; j++)
                {
                    worksheet.Range[i, j].Value = "Hello World!";
                }
            }

            workbook.Save(Stream.Null);
        }

        [Benchmark]
        public void ClosedXML()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                for (var i = 1; i <= ColumnNum; i++)
                {
                    for (var j = 1; j <= RowNum; j++)
                    {
                        worksheet.Cell(i, j).Value = "Hello World!";
                    }
                }

                workbook.SaveAs(Stream.Null);
            }
        }


        [Benchmark]
        public void OpenXml()
        {
            using (var spreadsheetDocument =
                SpreadsheetDocument.Create(Stream.Null, SpreadsheetDocumentType.Workbook))
            {
                var workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();

                var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                var sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                for (var i = 1; i <= ColumnNum; i++)
                {
                    for (var j = 1; j <= RowNum; j++)
                    {
                        var row = new Row();
                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue("Hello World!")
                        };
                        row.Append(cell);
                        sheetData.Append(row);
                    }
                }

                workbookpart.Workbook.Save();

                spreadsheetDocument.Close();
            }
        }

        [Benchmark]
        public void EPPlus()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                for (var i = 1; i <= ColumnNum; i++)
                {
                    for (var j = 1; j <= RowNum; j++)
                    {
                        worksheet.Cells[i, j].Value = "Hello World!";
                    }
                }
                package.SaveAs(Stream.Null);
            }
        }

        [Benchmark]
        public void NPOI()
        {
            var workbook = new XSSFWorkbook();
            var worksheet = workbook.CreateSheet("Sheet1");
            for (var i = 1; i <= ColumnNum; i++)
            {
                for (var j = 1; j <= RowNum; j++)
                {
                    worksheet.CreateRow(i).CreateCell(j).SetCellValue("Hello World!");
                }
            }
            workbook.Write(Stream.Null);
        }
    }
}
