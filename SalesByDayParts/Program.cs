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
            //program.SalesByDayPartsWeekly();

            program.SalesByDayPartsWeeklyNorth();
        }


        //Logical code for North
        public void SalesByDayPartsWeeklyNorth()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string SalesByDayParts = @"C:\Users\Nimap\Documents\Sales by day part\OneDrive_2024-03-01\Sales by DayParts\2024\Weekly\Sales By DayParts Week 7 CJNC.xlsx";

            string CJNorthXpientCy = @"C:\Users\Nimap\Documents\Sales by day part\OneDrive_2024-03-01\2024\2024\North\02-2024\02.20 - 02.26\2024-02-26.xlsx";

            string CJNorthXpientLy = @"C:\Users\Nimap\Documents\Sales by day part\2023\2023\North\02-2023\02.23 - 02.28-POSTED\2023-02-27.xlsx";

            string storeListFilePath = @"C:\Users\Public\Documents\StoreList.xlsx";

            Excel.Workbook storeList = excelApp.Workbooks.Open(storeListFilePath);

            Excel.Workbook SalesByDayPartsWorkbook = excelApp.Workbooks.Open(SalesByDayParts);

            Excel.Workbook CJNorthXpientCyWorkbook = excelApp.Workbooks.Open(CJNorthXpientCy);

            Excel.Workbook CJNorthXpientLyWorkbook = excelApp.Workbooks.Open(CJNorthXpientLy);

            try
            {
                //date by Week Code

                var date = "02/26/2024";

                DateTime dateValue;
                DateTime.TryParseExact(date, "MM/dd/yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out dateValue);
                Calendar cal = new CultureInfo("en-US").Calendar;
                cal = CultureInfo.CurrentCulture.Calendar;

                var currentWeekNbr = cal.GetWeekOfYear(dateValue, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                currentWeekNbr -= 1;

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

                // Adding the new Week

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

                //Adding the Ty sheet

                Worksheet tySummarySheet = CJNorthXpientCyWorkbook.Worksheets["Summary"];

                Worksheet salesTySheet = SalesByDayPartsWorkbook.Worksheets["TY"];

                int tySummaryLastRow = tySummarySheet.Cells[tySummarySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                int salesTyLastRow = salesTySheet.Cells[salesTySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                Excel.Range copyTySummaryRange = tySummarySheet.Range["A1:I" + tySummaryLastRow];
                copyTySummaryRange.Copy(Type.Missing);
                copyTySummaryRange.PasteSpecial(XlPasteType.xlPasteValues);
                copyTySummaryRange.Copy(Type.Missing);

                Excel.Range pasteTySummaryRange = salesTySheet.Range["A1:I" + salesTyLastRow];
                pasteTySummaryRange.PasteSpecial(XlPasteType.xlPasteAll);

                //Adding the Ly sheet

                Worksheet lySummarySheet = CJNorthXpientLyWorkbook.Worksheets["Summary"];

                Worksheet salesLySheet = SalesByDayPartsWorkbook.Worksheets["LY"];

                int lySummaryLastRow = lySummarySheet.Cells[lySummarySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                int salesLyLastRow = salesLySheet.Cells[salesLySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                Excel.Range copyLySummaryRange = lySummarySheet.Range["A1:I" + lySummaryLastRow];
                copyLySummaryRange.Copy(Type.Missing);
                copyLySummaryRange.PasteSpecial(XlPasteType.xlPasteValues);
                copyLySummaryRange.Copy(Type.Missing);

                Excel.Range pasteLySummaryRange = salesLySheet.Range["A1:I" + salesLyLastRow];
                pasteLySummaryRange.PasteSpecial(XlPasteType.xlPasteAll);

                

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






        //Logical code for South

        public void SalesByDayPartsWeekly()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string SalesByDayParts = @"C:\Users\Nimap\Documents\Sales by day part\OneDrive_2024-03-01\Sales by DayParts\2024\Weekly\Sales By DayParts Week 7 CJSC.xlsx";

            string CJSouthXpientCy = @"C:\Users\Nimap\Documents\Sales by day part\OneDrive_2024-03-01\2024\2024\South\02-2024\CJ South Xpient 2024-02-26.xlsx";

            string CJSouthXpientLy = @"C:\Users\Nimap\Documents\Sales by day part\2023\2023\South\02-2023\CJ South Xpient 2023-02-27.xlsx";

            string storeListFilePath = @"C:\Users\Public\Documents\StoreList.xlsx";

            string storeTimmimgPath = @"C:\Users\Nimap\Documents\Sales by day part\Sales By Dayparts\Final 02.21.23 - 02.27.23.xlsm";

            Excel.Workbook storeList = excelApp.Workbooks.Open(storeListFilePath);

            Excel.Workbook SalesByDayPartsWorkbook = excelApp.Workbooks.Open(SalesByDayParts);

            Excel.Workbook CJSouthXpientCyWorkbook = excelApp.Workbooks.Open(CJSouthXpientCy);

            Excel.Workbook CJSouthXpientLyWorkbook = excelApp.Workbooks.Open(CJSouthXpientLy);

            Excel.Workbook storeTimmimgWorkbook = excelApp.Workbooks.Open(storeTimmimgPath);

            try
            {
                //date by Week Code

                var date = "02/26/2024";

                DateTime dateValue;
                DateTime.TryParseExact(date, "MM/dd/yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out dateValue);
                Calendar cal = new CultureInfo("en-US").Calendar;
                cal = CultureInfo.CurrentCulture.Calendar;

                var currentWeekNbr = cal.GetWeekOfYear(dateValue, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                currentWeekNbr -= 1;

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

                // Adding the new Week

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

                //Adding the Ty sheet

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

                //Adding the Ly sheet

                Worksheet lySummarySheet = CJSouthXpientLyWorkbook.Worksheets["Summary"];

                Worksheet salesLySheet = SalesByDayPartsWorkbook.Worksheets["LY"];

                int lySummaryLastRow = lySummarySheet.Cells[lySummarySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                int salesLyLastRow = salesLySheet.Cells[salesLySheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                Excel.Range copyLySummaryRange = lySummarySheet.Range["A1:BK" + lySummaryLastRow];
                copyLySummaryRange.Copy(Type.Missing);
                copyLySummaryRange.PasteSpecial(XlPasteType.xlPasteValues);
                copyLySummaryRange.Copy(Type.Missing);

                Excel.Range pasteLySummaryRange = salesLySheet.Range["A1:BK" + salesLyLastRow];
                pasteLySummaryRange.PasteSpecial(XlPasteType.xlPasteAll);

                // Adding the Dynamic rows for new Week

                Worksheet storePriorSheet = SalesByDayPartsWorkbook.Worksheets["Store % of Prior"];

                int storePriorLastRow = storePriorSheet.Cells[storePriorSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 5;

                int StorePriorAboveRow = storePriorLastRow - 5;

                Excel.Range copyStorePriorRange = storePriorSheet.Range[$"A{StorePriorAboveRow}:BF{storePriorLastRow}"];
                copyStorePriorRange.Copy(Type.Missing);

                Excel.Range pasteStorePriorRange = storePriorSheet.Range[$"A{storePriorLastRow + 1}:BF{storePriorLastRow + 6}"];
                pasteStorePriorRange.PasteSpecial(XlPasteType.xlPasteAll);

                Excel.Range storePriorFormulaChangeRange1 = storePriorSheet.Range[$"D{storePriorLastRow + 1}:BF{storePriorLastRow + 1}"];

                storePriorFormulaChangeRange1.Formula = $"=XLOOKUP(D$2, 'Week {currentWeekNbr}'!$A:$A, 'Week {currentWeekNbr}'!$F:$F, 0, 0, 1)";

                Excel.Range storePriorFormulaChangeRange2 = storePriorSheet.Range[$"D{storePriorLastRow + 2}:BF{storePriorLastRow + 2}"];

                storePriorFormulaChangeRange2.Formula = $"=XLOOKUP(D$2, 'Week {currentWeekNbr}'!$A:$A, 'Week {currentWeekNbr}'!$K:$K, 0, 0, 1)";

                Excel.Range storePriorFormulaChangeRange3 = storePriorSheet.Range[$"D{storePriorLastRow + 3}:BF{storePriorLastRow + 3}"];

                storePriorFormulaChangeRange3.Formula = $"=XLOOKUP(D$2, 'Week {currentWeekNbr}'!$A:$A, 'Week {currentWeekNbr}'!$P:$P, 0, 0, 1)";

                Excel.Range storePriorFormulaChangeRange4 = storePriorSheet.Range[$"D{storePriorLastRow + 4}:BF{storePriorLastRow + 4}"];

                storePriorFormulaChangeRange4.Formula = $"=XLOOKUP(D$2, 'Week {currentWeekNbr}'!$A:$A, 'Week {currentWeekNbr}'!$U:$U, 0, 0, 1)";

                Excel.Range storePriorFormulaChangeRange5 = storePriorSheet.Range[$"D{storePriorLastRow + 5}:BF{storePriorLastRow + 5}"];

                storePriorFormulaChangeRange5.Formula = $"=XLOOKUP(D$2, 'Week {currentWeekNbr}'!$A:$A, 'Week {currentWeekNbr}'!$Z:$Z, 0, 0, 1)";

                Excel.Range storePriorFormulaChangeRange6 = storePriorSheet.Range[$"D{storePriorLastRow + 6}:BF{storePriorLastRow + 6}"];

                storePriorFormulaChangeRange6.Formula = $"=XLOOKUP(D$2, 'Week {currentWeekNbr}'!$A:$A, 'Week {currentWeekNbr}'!$AE:$AE, 0, 0, 1)";

                // SiteList Update

                Worksheet cjListing = SalesByDayPartsWorkbook.Worksheets["Site List"];
                Worksheet siteList = storeList.Worksheets[1];

                Excel.Range copySiteRange = siteList.Range["A1:N" + siteList.Rows.Count];
                copySiteRange.Copy(Type.Missing);

                Excel.Range pasteSiteRange = cjListing.Range["A1:N" + cjListing.Rows.Count];
                pasteSiteRange.PasteSpecial(XlPasteType.xlPasteAll);

                //StoreTimming update

                Worksheet salesStoreTimming = SalesByDayPartsWorkbook.Worksheets["Sheet2"];
                Worksheet finalStoreTimmimg = storeTimmimgWorkbook.Worksheets["Summary"];

                Excel.Range copyStoreRange = finalStoreTimmimg.Range["R1:V25"];
                copyStoreRange.Copy(Type.Missing);

                Excel.Range pasteStoreRange = salesStoreTimming.Range["A2:E25"];
                pasteStoreRange.PasteSpecial(XlPasteType.xlPasteAll);

                // Fetching weekly data 

                Worksheet newWeekSheet = SalesByDayPartsWorkbook.Worksheets[$"Week {currentWeekNbr}"];

                Excel.Range cellA1 = newWeekSheet.Cells[1, 1];
                Excel.Range cellA2 = newWeekSheet.Cells[1, 2];
                Excel.Range unmergedRange = newWeekSheet.Range[cellA1, cellA2];
                unmergedRange.UnMerge();
                cellA1.Clear();
                cellA2.Clear();

                unmergedRange.Merge();
                cellA1.MergeArea.Value = date;

                int weekDataRow = newWeekSheet.Cells[newWeekSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                Excel.Range weekDataRange = newWeekSheet.Range[$"C{weekDataRow}:AE{weekDataRow}"];
                weekDataRange.Copy(Type.Missing);

                Worksheet ytdSheet = SalesByDayPartsWorkbook.Worksheets["YTD-2023"];

                int ytdLastRow = ytdSheet.Cells[ytdSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                for (int i = 1; i < ytdLastRow; i++)
                {
                    string weekData = Convert.ToString(ytdSheet.Cells[i, 1].Value);

                    if (!string.IsNullOrEmpty(weekData) && weekData.Equals(Convert.ToString(currentWeekNbr)))
                    {
                        Excel.Range ytdDataRange = ytdSheet.Range[$"C{i}:AE{i}"];
                        ytdDataRange.PasteSpecial(XlPasteType.xlPasteValues);

                        break;

                    }
                }

                ytdSheet.Columns.AutoFit();

                // Completed Logical code for South part


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
