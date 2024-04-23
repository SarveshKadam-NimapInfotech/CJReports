using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Ink;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SalesTax
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.SalesTax();

        }

        public void SalesTax()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string salesTaxFilePath = @"C:\Users\Nimap\Documents\SalesTax\CA Sales Tax 1Qtr 2024-02 Excl EBT Sales.xlsx";

            string glFilePath = @"C:\Users\Nimap\Documents\SalesTax\All Live GL 2023-2024 updated.xlsx";

            string uberFilePath = @"C:\Users\Nimap\Documents\SalesTax\eb2d98fb-f0db-4e95-aebf-df389fe780cb-united_states.csv";

            Excel.Workbook salesTaxWorkbook = excelApp.Workbooks.Open(salesTaxFilePath);

            Excel.Workbook glWorkbook = excelApp.Workbooks.Open(glFilePath);

            Excel.Workbook uberWorkbook = excelApp.Workbooks.Open(uberFilePath);



            try
            {
                string date = "02/29/2024";
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


                Worksheet salesTaxSummarySheet = salesTaxWorkbook.Worksheets["Summary"];

                salesTaxSummarySheet.Range["B2"].Value = month;
                salesTaxSummarySheet.Range["B3"].Value = year;
                salesTaxSummarySheet.Range["B4"].Value = date;

                switch (monthInt)
                {
                    case 2:
                        if (previousMonth == 1)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"C8:C14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"D8:D14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 3:
                        if (previousMonth == 2)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"D8:C14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"E8:E14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 4:
                        if (previousMonth == 3)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"E8:E14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"F8:F14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 5:
                        if (previousMonth == 4)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"F8:F14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"G8:G14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 6:
                        if (previousMonth == 5)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"G8:G14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"H8:H14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 7:
                        if (previousMonth == 6)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"H8:H14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"I8:I14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 8:
                        if (previousMonth == 7)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"I8:I14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"J8:J14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 9:
                        if (previousMonth == 8)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"J8:J14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"K8:K14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 10:
                        if (previousMonth == 9)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"K8:K14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"L8:L14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 11:
                        if (previousMonth == 10)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"L8:L14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"M8:M14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 12:
                        if (previousMonth == 11)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"M8:M14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"N8:N14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    default:
                        break;
                }

                Worksheet glSheet = glWorkbook.Worksheets[1];


                var glSheetFilterListYear = new object[]
                {
                    year
                };

                var glSheetFilterListMonth = new object[]
                {
                    month
                };

                var glSheetFilterList1 = new object[]
                {
                    "Sales Tax Payable"
                };

                var glSheetFilterList2 = new object[]
                {
                    "Bank JE"

                };

                Range sourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                glSheet.ShowAllData();
                sourceRange.AutoFilter(3, glSheetFilterListYear, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterListMonth, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(10, glSheetFilterList1, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(13, glSheetFilterList2, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                //Range filteredRange = sourceRange.SpecialCells(XlCellType.xlCellTypeVisible);

                Range copyRange = glSheet.Range["A1:W" + glSheet.Rows.Count];

                Worksheet newSheet = glWorkbook.Worksheets.Add();

                Range pasteRange = newSheet.Range["A1:W" + newSheet.Rows.Count];

                copyRange.Copy(Type.Missing);
                pasteRange.PasteSpecial(XlPasteType.xlPasteAll);


                List<KeyValuePair<string, string>> dataList = new List<KeyValuePair<string, string>>();

                int newSheetLastRow = newSheet.Cells[newSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                for (int i = 2; i <= newSheetLastRow; i++)
                {
                    string key = newSheet.Cells[i, 2].Value2?.ToString();

                    string value = newSheet.Cells[i, 14].Value2?.ToString();

                    if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                    {
                        dataList.Add(new KeyValuePair<string, string>(key, value));
                    }
                }

                newSheet.Delete();

                //foreach (var pair in dataList)
                //{
                //    // Find the key in column B of the summary sheet
                //    Excel.Range keyRange = salesTaxSummarySheet.Range["B18:B24"].Find(pair.Key, Type.Missing,
                //        Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows,
                //        Excel.XlSearchDirection.xlNext, false, false, Type.Missing);

                //    // If the key is found
                //    if (keyRange != null)
                //    {
                //        // Get the row number where the key is found
                //        int keyRow = keyRange.Row;

                //        // Print the value in the corresponding row of column C
                //        salesTaxSummarySheet.Cells[keyRow, 3].Value2 = pair.Value;
                //    }
                //}

                switch (monthInt)
                {
                    case 2:
                       
                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 3].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }
                        
                        break;

                    case 3:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 4].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 4:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 5].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 5:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 6].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }


                        break;

                    case 6:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 7].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 7:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 8].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 8:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 9].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 9:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 10].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 10:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 11].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 11:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 12].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 12:

                        foreach (var pair in dataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B18:B24"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 13].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    default:
                        break;
                }
                
                //glSheet.AutoFilterMode = false;

                var glSheetFilterList3 = new object[]
                {
                    "Sales Tax Payable"
                };

                var glSheetFilterList4 = new object[]
                {
                    "Sales",
                    "Sales Refund-3rd Party-02.2024",
                    "Uber Sales Tax-02.2024"


                };

                //Range sourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                glSheet.ShowAllData();
                sourceRange.AutoFilter(3, glSheetFilterListYear, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterListMonth, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(10, glSheetFilterList3, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(13, glSheetFilterList4, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range copyRange2 = glSheet.Range["A1:W" + glSheet.Rows.Count];

                Worksheet newSheet1 = glWorkbook.Worksheets.Add();

                Range pasteRange2 = newSheet1.Range["A1:W" + newSheet1.Rows.Count];

                copyRange2.Copy(Type.Missing);
                pasteRange2.PasteSpecial(XlPasteType.xlPasteAll);

                Worksheet Pivot = newSheet1;
                Range pivotData = newSheet1.Range[$"A:W"];

                PivotTable pivotTable = newSheet1.PivotTableWizard(XlPivotTableSourceType.xlDatabase, pivotData, Pivot.Range["Z2"], "PIVOT");

                PivotField amountFields = pivotTable.PivotFields("Amount");
                amountFields.Orientation = XlPivotFieldOrientation.xlDataField;
                amountFields.Function = XlConsolidationFunction.xlSum;

                PivotField storeFields = pivotTable.PivotFields("Entity");
                storeFields.Orientation = XlPivotFieldOrientation.xlRowField;

                Range pivotCopyRange = newSheet1.Range["Z2:AA12"];
                Range pivotPasteRange = newSheet1.Range["Z14:AA24"];

                pivotCopyRange.Copy(Type.Missing);
                pivotPasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                List<KeyValuePair<string, string>> pivotDataList = new List<KeyValuePair<string, string>>();

                for (int i = 14; i <= 24; i++)
                {
                    string key = newSheet1.Cells[i, 26].Value2?.ToString();

                    string value = newSheet1.Cells[i, 27].Value2?.ToString();

                    if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                    {
                        pivotDataList.Add(new KeyValuePair<string, string>(key, value));
                    }
                }

                switch (monthInt)
                {
                    case 2:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        if (double.TryParse(pair.Value, out double doubleValue) ) // Assuming pair.Value is of type string
                                        {
                                            salesTaxSummarySheet.Cells[keyRow, 4].Value2 = doubleValue * -1;
                                        }

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 3:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 5].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 4:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 6].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 5:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 7].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }


                        break;

                    case 6:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 8].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 7:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 9].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 8:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 10].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 9:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 11].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 10:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 12].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 11:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 13].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    case 12:

                        foreach (var pair in pivotDataList)
                        {
                            if (pair.Key.Length >= 2)
                            {
                                string prefix = pair.Key.Substring(0, 2);

                                Excel.Range searchRange = salesTaxSummarySheet.Range["B38:B44"];

                                foreach (Excel.Range cell in searchRange)
                                {
                                    string cellValue = cell.Value2?.ToString();

                                    if (!string.IsNullOrEmpty(cellValue) && cellValue.StartsWith(prefix))
                                    {
                                        int keyRow = cell.Row;

                                        salesTaxSummarySheet.Cells[keyRow, 14 ].Value2 = pair.Value;

                                        break;
                                    }
                                }
                            }
                        }

                        break;

                    default:
                        break;
                }

                newSheet1.Delete();

                Worksheet salesData = salesTaxWorkbook.Worksheets["Sales Data as per P&L Net Sales"];

                switch (monthInt)
                {
                    case 2:
                        if (previousMonth == 1)
                        {
                            Range copyRange1 = salesData.Range[$"E4:E74"];
                            Range pasteRange1 = salesData.Range[$"F4:F74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break; 
                    
                    case 3:
                        if (previousMonth == 2)
                        {
                            Range copyRange1 = salesData.Range[$"F4:F74"];
                            Range pasteRange1 = salesData.Range[$"G4:G74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 4:
                        if (previousMonth == 3)
                        {
                            Range copyRange1 = salesData.Range[$"G4:G74"];
                            Range pasteRange1 = salesData.Range[$"H4:H74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 5:
                        if (previousMonth == 4)
                        {
                            Range copyRange1 = salesData.Range[$"H4:H74"];
                            Range pasteRange1 = salesData.Range[$"I4:I74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 6:
                        if (previousMonth == 5)
                        {
                            Range copyRange1 = salesData.Range[$"I4:I74"];
                            Range pasteRange1 = salesData.Range[$"J4:J74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 7:
                        if (previousMonth == 6)
                        {
                            Range copyRange1 = salesData.Range[$"J4:J74"];
                            Range pasteRange1 = salesData.Range[$"K4:K74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 8:
                        if (previousMonth == 7)
                        {
                            Range copyRange1 = salesData.Range[$"K4:K74"];
                            Range pasteRange1 = salesData.Range[$"L4:L74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 9:
                        if (previousMonth == 8)
                        {
                            Range copyRange1 = salesData.Range[$"L4:L74"];
                            Range pasteRange1 = salesData.Range[$"M4:M74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 10:
                        if (previousMonth == 9)
                        {
                            Range copyRange1 = salesData.Range[$"M4:M74"];
                            Range pasteRange1 = salesData.Range[$"N4:N74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 11:
                        if (previousMonth == 10)
                        {
                            Range copyRange1 = salesData.Range[$"N4:N74"];
                            Range pasteRange1 = salesData.Range[$"O4:O74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 12:
                        if (previousMonth == 11)
                        {
                            Range copyRange1 = salesData.Range[$"O4:O74"];
                            Range pasteRange1 = salesData.Range[$"P4:P74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    default:
                        break;
                }

                Worksheet salesGL = salesTaxWorkbook.Worksheets["Sales GL"];
                salesGL.ShowAllData();
                Range clearSalesSheet = salesGL.Range["A1:W" + glSheet.Rows.Count];
                clearSalesSheet.Clear();

                var glSheetFilterList5 = new object[]
                {
                    "Net Sales"
                };

                var glSheetFilterList6 = new object[]
                {
                    "CJ"
                };

                //Range sourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                glSheet.ShowAllData();
                sourceRange.AutoFilter(3, glSheetFilterListYear, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterListMonth, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(8, glSheetFilterList5, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(1, glSheetFilterList6, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range copyRange3 = glSheet.Range["A1:W" + glSheet.Rows.Count];

                Range pasteRange3 = salesGL.Range["A1:W" + salesGL.Rows.Count];

                copyRange3.Copy(Type.Missing);
                pasteRange3.PasteSpecial(XlPasteType.xlPasteAll);

                Worksheet ebtData = salesTaxWorkbook.Worksheets["EBT from 10023"];

                switch (monthInt)
                {
                    case 2:
                        if (previousMonth == 1)
                        {
                            Range copyRange1 = ebtData.Range[$"E4:E74"];
                            Range pasteRange1 = ebtData.Range[$"F4:F74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 3:
                        if (previousMonth == 2)
                        {
                            Range copyRange1 = ebtData.Range[$"F4:F74"];
                            Range pasteRange1 = ebtData.Range[$"G4:G74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 4:
                        if (previousMonth == 3)
                        {
                            Range copyRange1 = ebtData.Range[$"G4:G74"];
                            Range pasteRange1 = ebtData.Range[$"H4:H74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 5:
                        if (previousMonth == 4)
                        {
                            Range copyRange1 = ebtData.Range[$"H4:H74"];
                            Range pasteRange1 = ebtData.Range[$"I4:I74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 6:
                        if (previousMonth == 5)
                        {
                            Range copyRange1 = ebtData.Range[$"I4:I74"];
                            Range pasteRange1 = ebtData.Range[$"J4:J74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 7:
                        if (previousMonth == 6)
                        {
                            Range copyRange1 = ebtData.Range[$"J4:J74"];
                            Range pasteRange1 = ebtData.Range[$"K4:K74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 8:
                        if (previousMonth == 7)
                        {
                            Range copyRange1 = ebtData.Range[$"K4:K74"];
                            Range pasteRange1 = ebtData.Range[$"L4:L74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 9:
                        if (previousMonth == 8)
                        {
                            Range copyRange1 = ebtData.Range[$"L4:L74"];
                            Range pasteRange1 = ebtData.Range[$"M4:M74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 10:
                        if (previousMonth == 9)
                        {
                            Range copyRange1 = ebtData.Range[$"M4:M74"];
                            Range pasteRange1 = ebtData.Range[$"N4:N74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 11:
                        if (previousMonth == 10)
                        {
                            Range copyRange1 = ebtData.Range[$"N4:N74"];
                            Range pasteRange1 = ebtData.Range[$"O4:O74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 12:
                        if (previousMonth == 11)
                        {
                            Range copyRange1 = ebtData.Range[$"O4:O74"];
                            Range pasteRange1 = ebtData.Range[$"P4:P74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    default:
                        break;
                }

                Worksheet ebtSource = salesTaxWorkbook.Worksheets["EBT SOURCE"];

                Range ebtClearRange = ebtSource.Range["A2:G" + ebtSource.Rows.Count];
                ebtClearRange.Clear();

                var glSheetFilterList7 = new object[]
               {
                    "EBT food",
                    "EBT Cash"  
               };

                var modifiedFilterList = glSheetFilterList7.Select(item => $"*{item}*").ToArray();

                //Range sourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                glSheet.ShowAllData();
                sourceRange.AutoFilter(3, glSheetFilterListYear, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterListMonth, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(12, modifiedFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range copyEnity = glSheet.Range["B2:B" + glSheet.Rows.Count];
                Range pasteEnity = ebtSource.Range["A2:A" + ebtSource.Rows.Count];

                copyEnity.Copy(Type.Missing);
                pasteEnity.PasteSpecial(XlPasteType.xlPasteAll);

                Range copyPer = glSheet.Range["D2:D" + glSheet.Rows.Count];
                Range pastePer = ebtSource.Range["B2:B" + ebtSource.Rows.Count];

                copyPer.Copy(Type.Missing);
                pastePer.PasteSpecial(XlPasteType.xlPasteAll);

                Range copyDate = glSheet.Range["E2:E" + glSheet.Rows.Count];
                Range pasteDate = ebtSource.Range["C2:C" + ebtSource.Rows.Count];

                copyDate.Copy(Type.Missing);
                pasteDate.PasteSpecial(XlPasteType.xlPasteAll);

                Range copyJEAndComment = glSheet.Range["K2:L" + glSheet.Rows.Count];
                Range pasteJEAndComment = ebtSource.Range["D2:E" + ebtSource.Rows.Count];

                copyJEAndComment.Copy(Type.Missing);
                pasteJEAndComment.PasteSpecial(XlPasteType.xlPasteAll);

                Range copyDebitAndCredit = glSheet.Range["P2:Q" + glSheet.Rows.Count];
                Range pasteDebitAndCredit = ebtSource.Range["F2:G" + ebtSource.Rows.Count];

                copyDebitAndCredit.Copy(Type.Missing);
                pasteDebitAndCredit.PasteSpecial(XlPasteType.xlPasteAll);


                //List<glEbtData> glEbtDataList = new List<glEbtData>();

                //glSheet.ShowAllData();
                //int glLastRow = glSheet.Cells[glSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                //for(int i = 1; i <= glLastRow; i++)
                //{
                //    string entity = Convert.ToString(glSheet.Cells[i, 2].Value);
                //    string per = Convert.ToString(glSheet.Cells[i, 4].Value);
                //    string ebtDate = Convert.ToString(glSheet.Cells[i, 5].Value);
                //    string je = Convert.ToString(glSheet.Cells[i, 11].Value);
                //    string comment = Convert.ToString(glSheet.Cells[i, 12].Value);
                //    string debit = Convert.ToString(glSheet.Cells[i, 16].Value);
                //    string credit = Convert.ToString(glSheet.Cells[i, 17].Value);

                //    if (per.Equals("2") && per.Equals("2024"))
                //    {
                //        if (comment.Contains("EBT food") || comment.Contains("EBT Cash"))
                //        {
                //            glEbtData rowData = new glEbtData
                //            {
                //                Entity = entity,
                //                Per = per,
                //                EBTDate = ebtDate,
                //                JE = je,
                //                Comment = comment,
                //                Debit = debit,
                //                Credit = credit,


                //            };

                //            glEbtDataList.Add(rowData);
                //        }

                //    }
                //}

                //int ebtRowCounter = 2;
                //foreach (var data in glEbtDataList)
                //{
                //    ebtSource.Cells[ebtRowCounter, 1].Value = data.Entity;
                //    ebtSource.Cells[ebtRowCounter, 2].Value = data.Per;
                //    ebtSource.Cells[ebtRowCounter, 3].Value = data.EBTDate;
                //    ebtSource.Cells[ebtRowCounter, 4].Value = data.JE;
                //    ebtSource.Cells[ebtRowCounter, 5].Value = data.Comment;
                //    ebtSource.Cells[ebtRowCounter, 6].Value = data.Debit;
                //    ebtSource.Cells[ebtRowCounter, 7].Value = data.Credit;
                //    ebtRowCounter++;
                //}


                Worksheet uberData = salesTaxWorkbook.Worksheets["UBER"];

                switch (monthInt)
                {
                    case 2:
                        if (previousMonth == 1)
                        {
                            Range copyRange1 = uberData.Range[$"E4:E74"];
                            Range pasteRange1 = uberData.Range[$"F4:F74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 3:
                        if (previousMonth == 2)
                        {
                            Range copyRange1 = uberData.Range[$"F4:F74"];
                            Range pasteRange1 = uberData.Range[$"G4:G74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 4:
                        if (previousMonth == 3)
                        {
                            Range copyRange1 = uberData.Range[$"G4:G74"];
                            Range pasteRange1 = uberData.Range[$"H4:H74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 5:
                        if (previousMonth == 4)
                        {
                            Range copyRange1 = uberData.Range[$"H4:H74"];
                            Range pasteRange1 = uberData.Range[$"I4:I74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 6:
                        if (previousMonth == 5)
                        {
                            Range copyRange1 = uberData.Range[$"I4:I74"];
                            Range pasteRange1 = uberData.Range[$"J4:J74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 7:
                        if (previousMonth == 6)
                        {
                            Range copyRange1 = uberData.Range[$"J4:J74"];
                            Range pasteRange1 = uberData.Range[$"K4:K74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 8:
                        if (previousMonth == 7)
                        {
                            Range copyRange1 = uberData.Range[$"K4:K74"];
                            Range pasteRange1 = uberData.Range[$"L4:L74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 9:
                        if (previousMonth == 8)
                        {
                            Range copyRange1 = uberData.Range[$"L4:L74"];
                            Range pasteRange1 = uberData.Range[$"M4:M74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 10:
                        if (previousMonth == 9)
                        {
                            Range copyRange1 = uberData.Range[$"M4:M74"];
                            Range pasteRange1 = uberData.Range[$"N4:N74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 11:
                        if (previousMonth == 10)
                        {
                            Range copyRange1 = uberData.Range[$"N4:N74"];
                            Range pasteRange1 = uberData.Range[$"O4:O74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 12:
                        if (previousMonth == 11)
                        {
                            Range copyRange1 = uberData.Range[$"O4:O74"];
                            Range pasteRange1 = uberData.Range[$"P4:P74"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    default:
                        break;
                }

                Worksheet uberSheet = uberWorkbook.Worksheets[1];






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

    internal class glEbtData
    {
        public string Entity { get; set; }
        public string Per { get; set; }
        public string EBTDate { get; set; }
        public string JE { get; set; }
        public string Comment { get; set; }
        public string Debit { get; set; }
        public string Credit { get; set; }

    }
}
