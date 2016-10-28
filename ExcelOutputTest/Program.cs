using System;
using OfficeOpenXml;
using System.Diagnostics;
using System.Data;
using System.IO;

namespace ExcelOutputTest {
    class Program {

        private const string ExcelFileName = "test.xlsx";
        private const string SheetName = "test";

        static void Main(string[] args) {

            if (File.Exists(ExcelFileName)) {
                File.Delete(ExcelFileName);
            }

            DataTable outputData = CreateTestData();

            var s = new Stopwatch();
            s.Start();

            var outputFile = new FileInfo(ExcelFileName);
            using (var book = new ExcelPackage(outputFile))
            using (var sheet = book.Workbook.Worksheets.Add(SheetName)) {
                CreateExcel(outputData, book, sheet);
            }

            s.Stop();
            Console.WriteLine($"終了({ s.ElapsedMilliseconds.ToString("#,0") }ms)");

        }

        /// <summary>
        /// 200 * 30,000のデータ作成
        /// </summary>
        /// <returns></returns>
        static DataTable CreateTestData() {
            var dt = new DataTable();

            for (var i = 0; i < 200; i++) {
                dt.Columns.Add($"カラム{ i }");
            }

            for (var i = 0; i < 30000; i++) {
                var dr = dt.NewRow();
                for (var j = 0; j < dt.Columns.Count; j++) {
                    dr[j] = $"テストデータ{ i }_{ j }";
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        /// <summary>
        /// (200 * 30,000) + 200 = 6,000,200セルにデータをつっこむ
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="book"></param>
        /// <param name="sheet"></param>
        static void CreateExcel(DataTable dt, ExcelPackage book, ExcelWorksheet sheet) {
            //カラム名設定
            for (var x = 1; x <= dt.Columns.Count; x++) {
                sheet.Cells[1, x].Value = dt.Columns[x - 1].ColumnName;
            }

            //データ設定
            for (var y = 2; y <= dt.Rows.Count + 1; y++) {

                var row = dt.Rows[y - 2];

                for (var x = 1; x <= dt.Columns.Count; x++) {

                    sheet.Cells[y, x].Value = row[x - 1];

                }
            }
            //セル幅調整 割りとアテにならない
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
            book.Save();
        }
    }
}
