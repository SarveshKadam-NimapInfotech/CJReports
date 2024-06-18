using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace CJNCStandardFoodCost
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            var templatePath = @"C:\Users\Nimap\Documents\CjncStandardFoodCost\previousMonthCjLabourStandardFoodCost-1713186755085.xlsx";

            var pnlFilePath = @"C:\Users\Nimap\Documents\CjncStandardFoodCost\pnlFile-1713186755547.xlsx";

            var glFilePath = @"C:\Users\Nimap\Documents\CjncStandardFoodCost\allLiveGl-1713186755584.xlsx";

            var payrollPeriodFilePath = @"C:\Users\Public\Documents\Payroll periods (1).xlsx";

            Excel.Workbook templateWorkbook = excelApp.Workbooks.Open(templatePath);
            Excel.Workbook pnlWorkbook = excelApp.Workbooks.Open(pnlFilePath);
            Excel.Workbook glWorkbook = excelApp.Workbooks.Open(glFilePath);
            Excel.Workbook payrollPeriodWorkbook = excelApp.Workbooks.Open(payrollPeriodFilePath);

            try
            {
                string date = "03/29/2024";
                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                int monthInt = parsedDate.Month;

                int yearInt = parsedDate.Year;

                int previousYear = yearInt - 1;

                string month = Convert.ToString(parsedDate.Month);

                string year = Convert.ToString(parsedDate.Year);

                int previousMonth = 0;
                if (monthInt == 1)
                {
                    previousMonth = 12;
                }
                else
                {
                    previousMonth = monthInt - 1;

                }

                string previousMonthString = previousMonth.ToString();
                string previousYearString = previousYear.ToString();

                string monthName = parsedDate.ToString("MMM", CultureInfo.InvariantCulture);

                var currentMonth = parsedDate.Month;
                int currentDate = parsedDate.Day;
                int currentYear = parsedDate.Year;
                string formattedDate = parsedDate.ToString("MM-yyyy");

                int monthRowCount = 1;
                DateTime startDate = DateTime.MinValue;
                DateTime endDate = DateTime.MinValue;
                DateTime previousMonthEndDate = DateTime.MinValue;

                string startDateInString = startDate.ToString();
                string endDateInString = endDate.ToString();
                string previousMonthEndDateInString = previousMonthEndDate.ToString();

                Worksheet payRollSheet = payrollPeriodWorkbook.Worksheets[1];

                Range monthColumn = payRollSheet.Range["E1:E" + payRollSheet.UsedRange.Rows.Count];

                for (int i = 1; i <= monthColumn.Rows.Count; i++)
                {

                    if (monthColumn.Cells[i, 1].Value == null || monthColumn.Cells[i, 1].Value.ToString() == "")
                    {
                        continue;
                    }
                    else if (monthColumn.Cells[i, 1].Value.ToString() == formattedDate)
                    {
                        if (monthRowCount == 1)
                        {
                            startDateInString = payRollSheet.Cells[i, 1].Value.ToString();
                            startDate = DateTime.Parse(startDateInString);
                            if (monthColumn.Cells[i - 1, 1].Value.ToString() != "Month")
                            {
                                previousMonthEndDateInString = payRollSheet.Cells[i - 1, 2].Value.ToString();
                                previousMonthEndDate = DateTime.Parse(previousMonthEndDateInString);
                            }
                            monthRowCount++;
                        }
                        else
                        {
                            endDateInString = payRollSheet.Cells[i, 2].Value.ToString();
                            endDate = DateTime.Parse(endDateInString);
                        }
                    }
                    else if (monthRowCount > 1)
                    {
                        break;
                    }
                }

                string startDateMonthNumber = Convert.ToString(startDate.Month);
                string endDateMonthNumber = Convert.ToString(endDate.Month);

                Worksheet salesTaxSummarySheet = templateWorkbook.Worksheets["Labor Standard Variance"];

                // Insert a new column at column J
                Excel.Range columnJ = salesTaxSummarySheet.Columns["J"];
                columnJ.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                Excel.Range columnK = salesTaxSummarySheet.Columns["K"];
                Excel.Range newColumnJ = salesTaxSummarySheet.Columns["J"];

                columnK.Copy(Type.Missing);
                newColumnJ.PasteSpecial(XlPasteType.xlPasteAll);
                columnK.PasteSpecial(XlPasteType.xlPasteValues);

                Excel.Range formulaRange = salesTaxSummarySheet.Range["J7:J15"];
                formulaRange.Formula = "=+H7+I7";

                Excel.Range formulaRange1 = salesTaxSummarySheet.Range["J17"];
                formulaRange1.Formula = "=SUM(J7:J15)";

                Excel.Range formulaRange2 = salesTaxSummarySheet.Range["J18"];
                formulaRange2.Formula = "=+J17/C17";

                Excel.Range cellJ3 = salesTaxSummarySheet.Range["J3"];
                cellJ3.Value = date;

                Excel.Range cellD3 = salesTaxSummarySheet.Range["D3"];
                cellD3.Value = date;

                Excel.Range cellF2 = salesTaxSummarySheet.Range["F2"];
                cellF2.Value = startDateInString;

                Excel.Range cellE2 = salesTaxSummarySheet.Range["E2"];
                cellE2.Value = endDateInString;

                Excel.Range cellC3 = salesTaxSummarySheet.Range["C3"];
                string valueInC3 = startDate.ToString("M/d") + " to " + endDate.ToString("M/d");
                cellC3.Value = valueInC3;

                Worksheet pnlSheet = pnlWorkbook.Worksheets[1];
                Worksheet templatePnlSheet = templateWorkbook.Worksheets[$"P&L {year}-{previousMonthString}"];
                Worksheet templatePnlDataSheet = templateWorkbook.Worksheets["P&L Data"];

                Excel.Range pnlCopyRange = pnlSheet.Range["A1:Q" + pnlSheet.Rows.Count];
                Excel.Range pnlPasteRange = templatePnlSheet.Range["C1:S" + templatePnlSheet.Rows.Count];

                pnlCopyRange.Copy(Type.Missing);
                pnlPasteRange.PasteSpecial(XlPasteType.xlPasteAll);

                templatePnlSheet.Name = $"P&L {year}-{month}";

                var pnlHeaderString = Convert.ToInt32(pnlSheet.Cells[3, 1].Value);

                templatePnlDataSheet.Cells[3, 4].Value = pnlHeaderString;
                templatePnlDataSheet.Cells[6, 5].Value = $"{month} {year}";
                templatePnlDataSheet.Cells[6, 7].Value = $"{month} {previousYearString}";
                templatePnlDataSheet.Cells[6, 13].Value = $"{month} {year}";
                templatePnlDataSheet.Cells[6, 15].Value = $"{month} {previousYearString}";

                var pnlNetSalesFilter = new object[]
                {
                    "         Net Sales"
                };
                Range sourceRange = pnlSheet.Range[pnlSheet.Cells[1, 1], pnlSheet.Cells[1, pnlSheet.UsedRange.Column]];
                pnlSheet.ShowAllData();
                sourceRange.AutoFilter(3, pnlNetSalesFilter, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range pnlDataCopyRange = pnlSheet.Range["A1:Q" + pnlSheet.Rows.Count];

                Worksheet pnlDataNewSheet = pnlWorkbook.Worksheets.Add();

                Range pnlDataPasteRange = pnlDataNewSheet.Range["A1:Q" + pnlDataNewSheet.Rows.Count];

                pnlDataCopyRange.Copy(Type.Missing);
                pnlDataPasteRange.PasteSpecial(XlPasteType.xlPasteAll);

                Range pnlDataCopyRange1 = pnlDataNewSheet.Range["A2:Q" + pnlDataNewSheet.Rows.Count];
                Range pnlDataPasteRange1 = templatePnlDataSheet.Range["D7:T" + templatePnlDataSheet.Rows.Count];


                Worksheet glSheet = glWorkbook.Worksheets[1];
                Worksheet salesSheet = templateWorkbook.Worksheets["Sales"];

                var glSheetFilterListGroup = new object[]
                {
                    "CJ"
                };
                var glSheetFilterListEntity = new object[]
                {
                   "2SH",
                   "3SI"

                };
                var glSheetFilterListYear = new object[]
                {
                   year
                };
                var glSheetFilterListMonth = new object[]
                {
                   startDateMonthNumber,
                   endDateMonthNumber
                };

                var glSheetFilterUser = new object[]
                {
                   "Net Sales"
                };

                Range glSourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                glSheet.ShowAllData();
                sourceRange.AutoFilter(1, glSheetFilterListGroup, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(2, glSheetFilterListEntity, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(3, glSheetFilterListYear, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterListMonth, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(5, ">=" + startDateInString, Excel.XlAutoFilterOperator.xlAnd, "<=" + endDateInString);
                sourceRange.AutoFilter(8, glSheetFilterUser, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range glCopyRange = glSheet.Range["A1:W" + glSheet.Rows.Count];

                Range glPasteRange = salesSheet.Range["A1:W" + salesSheet.Rows.Count];

                glCopyRange.Copy(Type.Missing);
                glPasteRange.Clear();
                glPasteRange.PasteSpecial(XlPasteType.xlPasteAll);

                Worksheet salesPivotSheet = templateWorkbook.Worksheets["Pivot Sales"];
                int salesPivotSheetLastRow = salesPivotSheet.Cells[salesPivotSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                Range salesPivotsheetFormulaRange = salesPivotSheet.Range["C4:C" + salesPivotSheetLastRow];
                salesPivotsheetFormulaRange.Formula = "=B4";




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
