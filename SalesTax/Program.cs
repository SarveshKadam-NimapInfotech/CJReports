﻿using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
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

            Excel.Workbook salesTaxWorkbook = excelApp.Workbooks.Open(salesTaxFilePath);

            Excel.Workbook glWorkbook = excelApp.Workbooks.Open(glFilePath);

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


                var glSheetFilterList1 = new object[]
                {
                    year
                };

                var glSheetFilterList2 = new object[]
                {
                    month
                };

                var glSheetFilterList3 = new object[]
                {
                    "Sales Tax Payable"
                };

                var glSheetFilterList4 = new object[]
                {
                    "Bank JE"

                };

                Range sourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                glSheet.ShowAllData();
                sourceRange.AutoFilter(3, glSheetFilterList1, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterList2, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(10, glSheetFilterList3, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(13, glSheetFilterList4, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range filteredRange = sourceRange.SpecialCells(XlCellType.xlCellTypeVisible);

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

                //glSheet.AutoFilterMode = false;

                var glSheetFilterList5 = new object[]
                {
                    year
                };

                var glSheetFilterList6 = new object[]
                {
                    month
                };

                var glSheetFilterList7 = new object[]
                {
                    "Sales Tax Payable"
                };

                var glSheetFilterList8 = new object[]
                {
                    "Sales",
                    "Sales Refund-3rd Party-02.2024",
                    "Uber Sales Tax-02.2024"


                };

                //Range sourceRange = glSheet.Range[glSheet.Cells[1, 1], glSheet.Cells[1, glSheet.UsedRange.Column]];
                glSheet.ShowAllData();
                sourceRange.AutoFilter(3, glSheetFilterList5, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(4, glSheetFilterList6, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(10, glSheetFilterList7, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                sourceRange.AutoFilter(13, glSheetFilterList8, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                Range filteredRange1 = sourceRange.SpecialCells(XlCellType.xlCellTypeVisible);

                Range copyRange1 = glSheet.Range["A1:W" + glSheet.Rows.Count];

                Worksheet newSheet1 = glWorkbook.Worksheets.Add();

                Range pasteRange1 = newSheet.Range["A1:W" + newSheet.Rows.Count];

                copyRange1.Copy(Type.Missing);
                pasteRange1.PasteSpecial(XlPasteType.xlPasteAll);









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
