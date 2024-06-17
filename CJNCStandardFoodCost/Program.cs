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

                Excel.Range formulaRange = salesTaxSummarySheet.Columns["J7:J15"];
                formulaRange.Formula = "=+H7+I7";

                Excel.Range formulaRange1 = salesTaxSummarySheet.Columns["J17"];
                formulaRange1.Formula = "=SUM(J7:J15)";

                Excel.Range formulaRange2 = salesTaxSummarySheet.Columns["J18"];
                formulaRange1.Formula = "=+J17/C17";

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
