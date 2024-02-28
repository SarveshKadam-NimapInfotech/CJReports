using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace R_M
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.RNM();
        }

        public void RNM()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string RNMDetailsFilePath = @"C:\Users\Nimap\Documents\R&M\R & M Details 11.2023\R&M Details.xlsx";

            string GlFilePath = @"C:\Users\Nimap\Documents\R&M\GlFiles\GlFile12-2023.xlsx";

            Excel.Workbook RNMDetailsWorkbook = excelApp.Workbooks.Open(RNMDetailsFilePath);

            Excel.Workbook GlFileWorkbook = excelApp.Workbooks.Open(GlFilePath);

            try
            {
                string date = "12/15/2023";
                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                string month = Convert.ToString(parsedDate.Month);

                string year = Convert.ToString(parsedDate.Year);

                Worksheet glSheet = GlFileWorkbook.Worksheets[1];

                Worksheet rnmGlDetailsSheet = RNMDetailsWorkbook.Worksheets["GL Details"];

                Worksheet pivotSheet = RNMDetailsWorkbook.Worksheets["Pivot"];

                var glSheetFilterList1 = new object[]
                {
                    "RandM"
                };

                var glSheetFilterList5 = new object[]
                {
                    year
                };

                var glSheetFilterList6 = new object[]
                {
                    month
                };

                var glSheetFilterList2 = new object[]
                {
                    "CJ"
                };

                var glSheetFilterList3 = new object[]
                {
                    "DF2",
                    "HN2",
                    "SCL",
                    "SG2"

                };

                var glSheetFilterList4 = new object[]
                {
                    "Total Repair and Maintenance"

                };

                Range sourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                sourceRange.AutoFilter(7, glSheetFilterList1, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(3, glSheetFilterList5, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterList6, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(1, glSheetFilterList2, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(2, glSheetFilterList3, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(8, glSheetFilterList4, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range copyRange = glSheet.Range["A1:W" + glSheet.Rows.Count];
                Range pasteRange = rnmGlDetailsSheet.Range["A1:W" + rnmGlDetailsSheet.Rows.Count];

                copyRange.Copy(Type.Missing);
                pasteRange.PasteSpecial(XlPasteType.xlPasteAll);

                PivotTable pivotTable = pivotSheet.PivotTables(1);
                pivotTable.RefreshTable();

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); ;
            }
            finally
            {
                excelApp.Quit();
                Marshal.ReleaseComObject(excelApp);
            }
        }
    }
}
