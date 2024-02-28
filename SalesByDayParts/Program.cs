using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SalesByDayParts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.SalesByDayPartsWeekly();
        }

        public void SalesByDayPartsWeekly()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string SalesByDayParts = @"C:\Users\Nimap\Documents\Sales by day part\Sales By Dayparts\Weekly\Sales By DayParts Week 48 CJSC.xlsx";

            string CJSouthXpientCy = @"C:\Users\Nimap\Documents\Sales by day part\Sales By Dayparts\Weekly\CJ South Xpient 2023-11-27.xlsx";

            string CJSouthXpientLy = @"C:\Users\Nimap\Documents\Sales by day part\Sales By Dayparts\Weekly\CJ South Xpient 2022-11-28.xlsx";


            Excel.Workbook SalesByDayPartsWorkbook = excelApp.Workbooks.Open(SalesByDayParts);

            Excel.Workbook CJSouthXpientCyWorkbook = excelApp.Workbooks.Open(CJSouthXpientCy);

            Excel.Workbook CJSouthXpientLyWorkbook = excelApp.Workbooks.Open(CJSouthXpientLy);


            try
            {
                var date = "12/04/2023";

                DateTime dateValue;
                DateTime.TryParseExact(date, "MM/dd/yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out dateValue);
                Calendar cal = new CultureInfo("en-US").Calendar;
                cal = CultureInfo.CurrentCulture.Calendar;

                var currentWeekNbr = cal.GetWeekOfYear(dateValue, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                //currentWeekNbr -= 1;

                int previousWeekNbr;
                if (currentWeekNbr == 1)
                {
                    // If the current week number is 1, get the last week of the previous year
                    previousWeekNbr = cal.GetWeekOfYear(dateValue.AddYears(-1), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                }
                else
                {
                    // Otherwise, get the previous week of the current year
                    previousWeekNbr = currentWeekNbr - 1;
                }

                foreach (Excel.Worksheet worksheet in SalesByDayPartsWorkbook.Worksheets)
                {
                    string sheetName = worksheet.Name;
                    if (sheetName.Contains($"Week {previousWeekNbr}"))
                    {
                        worksheet.Copy(After: worksheet);

                        Excel.Worksheet copiedSheet = SalesByDayPartsWorkbook.Worksheets[worksheet.Index + 1];

                        copiedSheet.Name = $"Week {currentWeekNbr}";

                        Excel.Range range = worksheet.Cells;

                        range.Copy(Type.Missing);
                        range.PasteSpecial(XlPasteType.xlPasteValues);

                        break;
                    }
                }

                //Worksheet previousWeekSheet = SalesByDayPartsWorkbook.Worksheets[$"Week {previousWeekNbr}"];

                //Worksheet currentWeekSheet = SalesByDayPartsWorkbook.Worksheets[$"Week {currentWeekNbr}"];

                Worksheet tySummarySheet = CJSouthXpientCyWorkbook.Worksheets["Summary"];

                Worksheet salesTySheet = SalesByDayPartsWorkbook.Worksheets["TY"];

                int tySummaryLastRow = tySummarySheet.Cells[tySummarySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                int salesTyLastRow = salesTySheet.Cells[salesTySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;


                Excel.Range copyTySummaryRange = tySummarySheet.Range["A1:BK" + tySummaryLastRow];
                copyTySummaryRange.Copy(Type.Missing);
                copyTySummaryRange.PasteSpecial(XlPasteType.xlPasteValues);
                copyTySummaryRange.Copy(Type.Missing);

                Excel.Range pasteTySummaryRange = salesTySheet.Range["A1:BK" + salesTyLastRow];
                pasteTySummaryRange.PasteSpecial(XlPasteType.xlPasteAll);







                Worksheet storePriorSheet = SalesByDayPartsWorkbook.Worksheets["Store % of Prior"];

                int storePriorLastRow = storePriorSheet.Cells[storePriorSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 5;

                int StorePriorAboveRow = storePriorLastRow - 5;

                Excel.Range copyStorePriorRange = storePriorSheet.Range[$"A{StorePriorAboveRow}:BF{storePriorLastRow}"];
                copyStorePriorRange.Copy(Type.Missing);

                Excel.Range pasteStorePriorRange = storePriorSheet.Range[$"A{storePriorLastRow + 1}:BF{storePriorLastRow + 6}"];
                pasteStorePriorRange.PasteSpecial(XlPasteType.xlPasteAll);

                Excel.Range storePriorFormulaChangeRange = storePriorSheet.Range[$"D{storePriorLastRow + 1}:BF{storePriorLastRow + 6}"];

                storePriorFormulaChangeRange.Formula = $"XLOOKUP(AH$2, 'Week 48'!$A:$A, 'Week 48'!$F:$F, 0, 0, 1)";

                storePriorFormulaChangeRange.Formula = $"XLOOKUP(AH$2, 'Week {currentWeekNbr}'!$A:$A, 'Week {currentWeekNbr}'!$F:$F, 0, 0, 1)";















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
