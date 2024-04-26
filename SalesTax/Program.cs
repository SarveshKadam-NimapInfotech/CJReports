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
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Drawing.Chart;

namespace SalesTax
{
    internal class RowDataForUberReport
    {
        public dynamic StoreName { get; set; }
        public dynamic StoreID { get; set; }
        public dynamic OrderID { get; set; }
        public dynamic WorkflowID { get; set; }
        public dynamic DiningMode { get; set; }
        public dynamic PaymentMode { get; set; }
        public dynamic OrderChannel { get; set; }
        public dynamic OrderStatus { get; set; }
        public dynamic OrderDate { get; set; }
        public dynamic OrderAcceptingTime { get; set; }
        public dynamic CustomerStatus { get; set; }
        public dynamic SalesExc { get; set; }
        public dynamic TaxOnSales { get; set; }
        public dynamic SalesInc { get; set; }
        public dynamic RefundsExc { get; set; }
        public dynamic TaxOnRefund { get; set; }
        public dynamic RefundsInc { get; set; }
        public dynamic PriceAdjusExc { get; set; }
        public dynamic TaxOnPriceAdjus { get; set; }
        public dynamic PromotionOnItems { get; set; }
        public dynamic TaxOnPromotions { get; set; }
        public dynamic PromotionsOnDelivery { get; set; }
        public dynamic TaxOnPromotionsDelivery { get; set; }
        public dynamic BagFee { get; set; }
        public dynamic MarketingAdjus { get; set; }
        public dynamic TotalSales { get; set; }
        public dynamic MarketplaceFee { get; set; }
        public dynamic MarketplaceFeePer { get; set; }
        public dynamic DeliveryNetworkFee { get; set; }
        public dynamic OrderProcessingFee { get; set; }
        public dynamic MerchantFee { get; set; }
        public dynamic TaxOnMerchantFee { get; set; }
        public dynamic Tips { get; set; }
        public dynamic OtherPaymentsDesc { get; set; }
        public dynamic OtherPayments { get; set; }
        public dynamic MarketPlaceFaciliatorTax { get; set; }
        public dynamic BackupWithHoldingTax { get; set; }
        public dynamic TotalPayout { get; set; }
        public dynamic PayoutDate { get; set; }
        public dynamic MarkupAmount { get; set; }
        public dynamic MarkupTax { get; set; }
        public dynamic RetailerLoyaltyID { get; set; }
        public dynamic PayoutReferenceID { get; set; }
    }
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

            string uberFilePath = $"{Path.GetFileNameWithoutExtension(glFilePath)}.xlsx";
            string targetUberFilePath= @"C:\Users\Nimap\Documents\SalesTax\eb2d98fb-f0db-4e95-aebf-df389fe780cb-united_states.xlsx";

            string salesRefundFilePath = @"C:\Users\Nimap\Documents\SalesTax\Sales Refund 02.2024.xlsx";

            Excel.Workbook salesTaxWorkbook = excelApp.Workbooks.Open(salesTaxFilePath);

            Excel.Workbook glWorkbook = excelApp.Workbooks.Open(glFilePath);

            Excel.Workbook salesRefundWorkbook = excelApp.Workbooks.Open(salesRefundFilePath);



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

                
                Worksheet uberSummarySheet = salesRefundWorkbook.Worksheets["Uber Summary"];
                Worksheet salesRefundSheet = salesTaxWorkbook.Worksheets["Sales Refunds"];

                Range salesRefundClearRange = salesRefundSheet.Range["A2:C" + salesRefundSheet.Rows.Count];
                salesRefundClearRange.Clear();

                int breakRow = 66;
                for (int i = 5; i <= uberSummarySheet.UsedRange.Rows.Count; i++)
                {
                    var uberValue = uberSummarySheet.Cells[i, 8].Value;
                    if ((uberValue == null) || !(uberValue is string))
                    {
                        continue;
                    }
                    string grandtotal = Convert.ToString(uberSummarySheet.Cells[i, 8].Value);
                    if (grandtotal.ToLower().Trim().Contains("grand total"))
                    {
                        breakRow = i-1;
                        break;
                    }
                }

                Range copyUberSummary1 = uberSummarySheet.Range["H5:I" + breakRow];
                Range copyUberSummary2 = uberSummarySheet.Range["M5:M" + breakRow];
                Range pasteUberSummary1 = salesRefundSheet.Range["A2:B" + breakRow];
                Range pasteUberSummary2 = salesRefundSheet.Range["C2:C" + breakRow];

                copyUberSummary1.Copy(Type.Missing);
                pasteUberSummary1.PasteSpecial(XlPasteType.xlPasteValues);
                copyUberSummary2.Copy(Type.Missing);
                pasteUberSummary2.PasteSpecial(XlPasteType.xlPasteValues);

                salesTaxWorkbook.Save();
                salesTaxWorkbook.Close();
                glWorkbook.Close();
                salesRefundWorkbook.Close();
                Marshal.ReleaseComObject(salesTaxWorkbook);
                Marshal.ReleaseComObject(glWorkbook);
                Marshal.ReleaseComObject(salesRefundWorkbook);


                Excel.Workbook uberWorkbook = excelApp.Workbooks.Open(uberFilePath);
                uberWorkbook.SaveAs(targetUberFilePath, FileFormat: XlFileFormat.xlOpenXMLWorkbook);
                uberWorkbook.Close();
                Marshal.ReleaseComObject(uberWorkbook);

                List<RowDataForUberReport> uberEPData = new List<RowDataForUberReport>();

                using (var SalesTaxPackage = new ExcelPackage(new FileInfo(salesTaxFilePath)))
                using (var excelPackage = new ExcelPackage(new FileInfo(targetUberFilePath)))
                {
                    
                    var worksheet = excelPackage.Workbook.Worksheets[0];

                    int numRows = worksheet.Dimension.End.Row;
                    int numCols = worksheet.Dimension.End.Column;
                    int orderDateCol = 0;
                    Dictionary<string, int> columnsMap = new Dictionary<string, int>
                    {
                        { "storeName",1},
                        { "storeId",1},
                        { "orderId",1},
                        { "workflowId",1},
                        { "diningMode",1},
                        { "paymentMode",1},
                        { "orderChannel",1},
                        { "orderStatus",1},
                        { "orderDate",1},
                        { "orderAccept",1},
                        { "customerUber",1},
                        { "salesExcl",1},
                        { "taxOnSales",1},
                        { "salesInc",1},
                        { "refundsExcl",1},
                        { "taxOnRefund",1},
                        { "refundInc",1},
                        { "priceAdjExcl",1},
                        { "taxOnPrice",1},
                        { "promotionOnItem",1},
                        { "taxOnPromotion",1},
                        { "promotionOnDelivery",1},
                        { "taxOnPromotionDelivery",1},
                        { "backFee",1},
                        { "marketingAdj",1},
                        { "totalSales",1},
                        { "marketplaceFee",1},
                        { "marketplacePer",1},
                        { "deliveryNetwork",1},
                        { "orderProcessing",1},
                        { "merchantFee",1},
                        { "taxOnMerchantFee",1},
                        { "tips",1},
                        { "otherPaymentDesc",1},
                        { "otherPayments",1},
                        { "marketplaceFacilitate",1},
                        { "backupWithHolding",1},
                        { "totalPayout",1},
                        { "payoutDate",1},
                        { "markupAmount",1},
                        { "markupText",1},
                        { "retailerLoyalty",1},
                        { "payoutReference",1},
                    };
                    for(int col = 1; col <= numCols; col++)
                    {
                        string textCol = worksheet.Cells[2, col].Text.ToLower().Trim();
                        if(textCol.Contains("store name"))
                        {
                            columnsMap["storeName"] = col;
                        }else if(textCol.Contains("store id"))
                        {
                            columnsMap["storeId"] = col;
                        }
                        else if (textCol.Contains("order id"))
                        {
                            columnsMap["orderId"] = col;
                        }
                        else if (textCol.Contains("workflow"))
                        {
                            columnsMap["workflowId"] = col;
                        }
                        else if (textCol.Contains("dining"))
                        {
                            columnsMap["diningMode"] = col;
                        }
                        else if (textCol.StartsWith("payment"))
                        {
                            columnsMap["paymentMode"] = col;
                        }
                        else if (textCol.Contains("order channel"))
                        {
                            columnsMap["orderChannel"] = col;
                        }
                        else if (textCol.Contains("order status"))
                        {
                            columnsMap["orderStatus"] = col;
                        }
                        else if (textCol.Contains("order date"))
                        {
                            columnsMap["orderDate"] = col;
                        }
                        else if (textCol.Contains("order accept"))
                        {
                            columnsMap["orderAccept"] = col;
                        }
                        else if (textCol.Contains("customer uber"))
                        {
                            columnsMap["customerUber"] = col;
                        }
                        else if (textCol.Contains("sales (excl"))
                        {
                            columnsMap["salesExcl"] = col;
                        }
                        else if (textCol.Contains("tax on sales"))
                        {
                            columnsMap["taxOnSales"] = col;
                        }
                        else if (textCol.Contains("sales (inc"))
                        {
                            columnsMap["salesInc"] = col;
                        }
                        else if (textCol.Contains("refunds (exc"))
                        {
                            columnsMap["refundsExcl"] = col;
                        }
                        else if (textCol.Contains("tax on refund"))
                        {
                            columnsMap["taxOnRefund"] = col;
                        }
                        else if (textCol.Contains("refunds (inc"))
                        {
                            columnsMap["refundInc"] = col;
                        }
                        else if (textCol.StartsWith("price adj"))
                        {
                            columnsMap["priceAdjExcl"] = col;
                        }
                        else if (textCol.Contains("tax on price adj"))
                        {
                            columnsMap["taxOnPrice"] = col;
                        }
                        else if (textCol.Contains("promotions on items"))
                        {
                            columnsMap["promotionOnItem"] = col;
                        }
                        else if (textCol.Contains("tax on promotion on items"))
                        {
                            columnsMap["taxOnPromotion"] = col;
                        }
                        else if (textCol.StartsWith("promotions on delivery"))
                        {
                            columnsMap["promotionOnDelivery"] = col;
                        }
                        else if (textCol.StartsWith("tax on promotions on delivery"))
                        {
                            columnsMap["taxOnPromotionDelivery"] = col;
                        }
                        else if (textCol.StartsWith("bag"))
                        {
                            columnsMap["backFee"] = col;
                        }
                        else if (textCol.Contains("marketing adj"))
                        {
                            columnsMap["marketingAdj"] = col;
                        }
                        else if (textCol.Contains("total sales"))
                        {
                            columnsMap["totalSales"] = col;
                        }
                        else if (textCol.Equals("marketplace fee"))
                        {
                            columnsMap["marketplaceFee"] = col;
                        }
                        else if (textCol.Equals("marketplace fee %"))
                        {
                            columnsMap["marketplacePer"] = col;
                        }
                        else if (textCol.Contains("delivery network"))
                        {
                            columnsMap["deliveryNetwork"] = col;
                        }
                        else if (textCol.Contains("order processing fee"))
                        {
                            columnsMap["orderProcessing"] = col;
                        }
                        else if (textCol.StartsWith("merchant fee"))
                        {
                            columnsMap["merchantFee"] = col;
                        }
                        else if (textCol.Contains("tax on merchant fee"))
                        {
                            columnsMap["taxOnMerchantFee"] = col;
                        }
                        else if (textCol.Contains("tips"))
                        {
                            columnsMap["tips"] = col;
                        }
                        else if (textCol.Contains("other payments desc"))
                        {
                            columnsMap["otherPaymentDesc"] = col;
                        }
                        else if (textCol.Contains("other payments"))
                        {
                            columnsMap["otherPayments"] = col;
                        }
                        else if (textCol.Contains("marketplace faci"))
                        {
                            columnsMap["marketplaceFacilitate"] = col;
                        }
                        else if (textCol.Contains("backup withholding"))
                        {
                            columnsMap["backupWithHolding"] = col;
                        }
                        else if (textCol.Contains("total payout"))
                        {
                            columnsMap["totalPayout"] = col;
                        }
                        else if (textCol.Contains("payout date"))
                        {
                            columnsMap["payoutDate"] = col;
                        }
                        else if (textCol.Contains("markup amount"))
                        {
                            columnsMap["markupAmount"] = col;
                        }
                        else if (textCol.Contains("markup tax"))
                        {
                            columnsMap["markupText"] = col;
                        }
                        else if (textCol.Contains("retailer loyalty"))
                        {
                            columnsMap["retailerLoyalty"] = col;
                        }
                        else if (textCol.Contains("payout reference"))
                        {
                            columnsMap["payoutReference"] = col;
                        }


                    };
                    int intMonth = 2;
                    //List<RowDataForUberReport> uberDate = new List<RowDataForUberReport>();

                    for (int row = 3; row <= numRows; row++)
                    {
                        var value = worksheet.Cells[row, columnsMap["orderDate"]].Value;
                        DateTime orderDate = default;
                        if (value is DateTime valueDate)
                        {
                            orderDate = valueDate.Date;
                        } else if (value is double valueDouble)
                        {
                            orderDate = DateTime.FromOADate(valueDouble).Date;
                        }
                        if (orderDate.Equals(default) || orderDate.Month != intMonth)
                        {
                            continue;
                        }
                        RowDataForUberReport rowData = new RowDataForUberReport
                        {
                            StoreName = worksheet.Cells[row, columnsMap["storeName"]].Value,
                            StoreID = worksheet.Cells[row, columnsMap["storeId"]].Value,
                            OrderID = worksheet.Cells[row, columnsMap["orderId"]].Value,
                            WorkflowID = worksheet.Cells[row, columnsMap["workflowId"]].Value,
                            DiningMode = worksheet.Cells[row, columnsMap["diningMode"]].Value,
                            PaymentMode = worksheet.Cells[row, columnsMap["paymentMode"]].Value,
                            OrderChannel = worksheet.Cells[row, columnsMap["orderChannel"]].Value,
                            OrderStatus = worksheet.Cells[row, columnsMap["orderStatus"]].Value,
                            OrderDate = worksheet.Cells[row, columnsMap["orderDate"]].Value,
                            OrderAcceptingTime = worksheet.Cells[row, columnsMap["orderAccept"]].Value,
                            CustomerStatus = worksheet.Cells[row, columnsMap["customerUber"]].Value,
                            SalesExc = worksheet.Cells[row, columnsMap["salesExcl"]].Value,
                            TaxOnSales = worksheet.Cells[row, columnsMap["taxOnSales"]].Value,
                            SalesInc = worksheet.Cells[row, columnsMap["salesInc"]].Value,
                            RefundsExc = worksheet.Cells[row, columnsMap["refundsExcl"]].Value,
                            TaxOnRefund = worksheet.Cells[row, columnsMap["taxOnRefund"]].Value,
                            RefundsInc = worksheet.Cells[row, columnsMap["refundInc"]].Value,
                            PriceAdjusExc = worksheet.Cells[row, columnsMap["priceAdjExcl"]].Value,
                            TaxOnPriceAdjus = worksheet.Cells[row, columnsMap["taxOnPrice"]].Value,
                            PromotionOnItems = worksheet.Cells[row, columnsMap["promotionOnItem"]].Value,
                            TaxOnPromotions = worksheet.Cells[row, columnsMap["taxOnPromotion"]].Value,
                            PromotionsOnDelivery = worksheet.Cells[row, columnsMap["promotionOnDelivery"]].Value,
                            TaxOnPromotionsDelivery = worksheet.Cells[row, columnsMap["taxOnPromotionDelivery"]].Value,
                            BagFee = worksheet.Cells[row, columnsMap["backFee"]].Value,
                            MarketingAdjus = worksheet.Cells[row, columnsMap["marketingAdj"]].Value,
                            TotalSales = worksheet.Cells[row, columnsMap["totalSales"]].Value,
                            MarketplaceFee = worksheet.Cells[row, columnsMap["marketplaceFee"]].Value,
                            MarketplaceFeePer = worksheet.Cells[row, columnsMap["marketplacePer"]].Value,
                            DeliveryNetworkFee = worksheet.Cells[row, columnsMap["deliveryNetwork"]].Value,
                            OrderProcessingFee = worksheet.Cells[row, columnsMap["orderProcessing"]].Value,
                            MerchantFee = worksheet.Cells[row, columnsMap["merchantFee"]].Value,
                            TaxOnMerchantFee = worksheet.Cells[row, columnsMap["taxOnMerchantFee"]].Value,
                            Tips = worksheet.Cells[row, columnsMap["tips"]].Value,
                            OtherPaymentsDesc = worksheet.Cells[row, columnsMap["otherPaymentDesc"]].Value,
                            OtherPayments = worksheet.Cells[row, columnsMap["otherPayments"]].Value,
                            MarketPlaceFaciliatorTax = worksheet.Cells[row, columnsMap["marketplaceFacilitate"]].Value,
                            BackupWithHoldingTax = worksheet.Cells[row, columnsMap["backupWithHolding"]].Value,
                            TotalPayout = worksheet.Cells[row, columnsMap["totalPayout"]].Value,
                            PayoutDate = worksheet.Cells[row, columnsMap["payoutDate"]].Value,
                            MarkupAmount = worksheet.Cells[row, columnsMap["markupAmount"]].Value,
                            MarkupTax = worksheet.Cells[row, columnsMap["markupText"]].Value,
                            RetailerLoyaltyID = worksheet.Cells[row, columnsMap["retailerLoyalty"]].Value,
                            PayoutReferenceID = worksheet.Cells[row, columnsMap["payoutReference"]].Value,
                        };
                    uberEPData.Add(rowData);
                    }

                    var uberSourceSheet = SalesTaxPackage.Workbook.Worksheets["Uber source "];

                    uberSourceSheet.Cells["A2:AQ" + uberSourceSheet.Dimension.End.Row].Clear();

                    int uberRowCounter = 2;
                    foreach (var data in uberEPData)
                    {
                        uberSourceSheet.Cells[uberRowCounter, 1].Value = data.StoreName;
                        uberSourceSheet.Cells[uberRowCounter, 2].Value = data.StoreID;
                        uberSourceSheet.Cells[uberRowCounter, 3].Value = data.OrderID;
                        uberSourceSheet.Cells[uberRowCounter, 4].Value = data.WorkflowID;
                        uberSourceSheet.Cells[uberRowCounter, 5].Value = data.DiningMode;
                        uberSourceSheet.Cells[uberRowCounter, 6].Value = data.PaymentMode;
                        uberSourceSheet.Cells[uberRowCounter, 7].Value = data.OrderChannel;
                        uberSourceSheet.Cells[uberRowCounter, 8].Value = data.OrderStatus;
                        uberSourceSheet.Cells[uberRowCounter, 9].Value = data.OrderDate;
                        uberSourceSheet.Cells[uberRowCounter, 10].Value = data.OrderAcceptingTime;
                        uberSourceSheet.Cells[uberRowCounter, 11].Value = data.CustomerStatus;
                        uberSourceSheet.Cells[uberRowCounter, 12].Value = data.SalesExc;
                        uberSourceSheet.Cells[uberRowCounter, 13].Value = data.TaxOnSales;
                        uberSourceSheet.Cells[uberRowCounter, 14].Value = data.SalesInc;
                        uberSourceSheet.Cells[uberRowCounter, 15].Value = data.RefundsExc;
                        uberSourceSheet.Cells[uberRowCounter, 16].Value = data.TaxOnRefund;
                        uberSourceSheet.Cells[uberRowCounter, 17].Value = data.RefundsInc;
                        uberSourceSheet.Cells[uberRowCounter, 18].Value = data.PriceAdjusExc;
                        uberSourceSheet.Cells[uberRowCounter, 19].Value = data.TaxOnPriceAdjus;
                        uberSourceSheet.Cells[uberRowCounter, 20].Value = data.PromotionOnItems;
                        uberSourceSheet.Cells[uberRowCounter, 21].Value = data.TaxOnPromotions;
                        uberSourceSheet.Cells[uberRowCounter, 22].Value = data.PromotionsOnDelivery;
                        uberSourceSheet.Cells[uberRowCounter, 23].Value = data.TaxOnPromotionsDelivery;
                        uberSourceSheet.Cells[uberRowCounter, 24].Value = data.BagFee;
                        uberSourceSheet.Cells[uberRowCounter, 25].Value = data.MarketingAdjus;
                        uberSourceSheet.Cells[uberRowCounter, 26].Value = data.TotalSales;
                        uberSourceSheet.Cells[uberRowCounter, 27].Value = data.MarketplaceFee;
                        uberSourceSheet.Cells[uberRowCounter, 28].Value = data.MarketplaceFeePer;
                        uberSourceSheet.Cells[uberRowCounter, 29].Value = data.DeliveryNetworkFee;
                        uberSourceSheet.Cells[uberRowCounter, 30].Value = data.OrderProcessingFee;
                        uberSourceSheet.Cells[uberRowCounter, 31].Value = data.MerchantFee;
                        uberSourceSheet.Cells[uberRowCounter, 32].Value = data.TaxOnMerchantFee;
                        uberSourceSheet.Cells[uberRowCounter, 33].Value = data.Tips;
                        uberSourceSheet.Cells[uberRowCounter, 34].Value = data.OtherPaymentsDesc;
                        uberSourceSheet.Cells[uberRowCounter, 35].Value = data.OtherPayments;
                        uberSourceSheet.Cells[uberRowCounter, 36].Value = data.MarketPlaceFaciliatorTax;
                        uberSourceSheet.Cells[uberRowCounter, 37].Value = data.BackupWithHoldingTax;
                        uberSourceSheet.Cells[uberRowCounter, 38].Value = data.TotalPayout;
                        uberSourceSheet.Cells[uberRowCounter, 39].Value = data.PayoutDate;
                        uberSourceSheet.Cells[uberRowCounter, 40].Value = data.MarkupAmount;
                        uberSourceSheet.Cells[uberRowCounter, 41].Value = data.MarkupTax;
                        uberSourceSheet.Cells[uberRowCounter, 42].Value = data.RetailerLoyaltyID;
                        uberSourceSheet.Cells[uberRowCounter, 43].Value = data.PayoutReferenceID;

                        // Increment the row counter for the next iteration
                        uberRowCounter++;
                    }

                    SalesTaxPackage.Save();
                }

                




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

    //internal class glEbtData
    //{
    //    public string Entity { get; set; }
    //    public string Per { get; set; }
    //    public string EBTDate { get; set; }
    //    public string JE { get; set; }
    //    public string Comment { get; set; }
    //    public string Debit { get; set; }
    //    public string Credit { get; set; }

    //}
}
