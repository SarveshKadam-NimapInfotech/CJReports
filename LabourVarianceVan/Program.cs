using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace LabourVarianceVan
{
    internal class FfcglData
    {
        public string Entity { get; set; }
        public string Store { get; set; }
        public string JobDesp { get; set; }
        public string Desc { get; set; }
        public string Amount { get; set; }
    }

    internal class CjLabourStandardData
    {
        public string CjStore { get; set; }
        public double Sales { get; set; }
        public string Labour { get; set; }
        public string Actual { get; set; }
        public string Var { get; set; }
    }
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.LabourVarianceVan();
        }

         public void LabourVarianceVan()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string storeListFilePath = @"C:\Users\Public\Documents\StoreList.xlsx";
            Excel.Workbook storeList = excelApp.Workbooks.Open(storeListFilePath);

            string labourVar1FilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\Labor Var 2023-11-20.xlsx";
            Excel.Workbook labourVar1Workbook = excelApp.Workbooks.Open(labourVar1FilePath);

            string labourVar2FilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\Labor Var 2023.11.20.xlsx";
            Excel.Workbook labourVar2Workbook = excelApp.Workbooks.Open(labourVar2FilePath);

            string ffcglFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\FFCGL.xlsx";
            Excel.Workbook ffcglWorkbook = excelApp.Workbooks.Open(ffcglFilePath);

            string cjLabourFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\CJ Labor Standard 2023-11-20.xlsx";
            Excel.Workbook cjLabourWorkbook = excelApp.Workbooks.Open(cjLabourFilePath);

            string labourVarTrendFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\Labor Var Trend since 8.29.22.xlsx";
            Excel.Workbook labourVarTrendWorkbook = excelApp.Workbooks.Open(labourVarTrendFilePath);

            string weeklySalesFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\week 49\Week 49 Sales 2023-12-04.xlsm";
            Excel.Workbook weeklySalesWorkbook = excelApp.Workbooks.Open(weeklySalesFilePath);

            try
            {
                //first file - labour var van 1

                var date = "12/04/2023";

                DateTime dateValue;
                DateTime.TryParseExact(date, "MM/dd/yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out dateValue);
                Calendar cal = new CultureInfo("en-US").Calendar;
                cal = CultureInfo.CurrentCulture.Calendar;

                var currentWeekNbr = cal.GetWeekOfYear(dateValue, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

                var previousWeekNbr = currentWeekNbr - 1;

                
                foreach (Excel.Worksheet worksheet in weeklySalesWorkbook.Worksheets)
                {
                    string sheetName = worksheet.Name;
                    if (sheetName.StartsWith("Week"))
                    {
                        int weekNbr;
                        if (int.TryParse(sheetName.Split(' ')[1], out weekNbr))
                        {
                            if (weekNbr == currentWeekNbr || weekNbr == previousWeekNbr)
                            {
                                Excel.Worksheet lastWeekSheet = null;

                                foreach (Excel.Worksheet sheet in labourVar1Workbook.Worksheets)
                                {
                                    if (sheet.Name.StartsWith("Week "))
                                    {
                                        lastWeekSheet = sheet;
                                    }
                                }

                                worksheet.Copy(After: lastWeekSheet);
                            }
                        }
                    }
                }

                string GetColumnLetter(int columnNumber)
                {
                    int dividend = columnNumber;
                    string columnLetter = string.Empty;
                    int modulo;

                    while (dividend > 0)
                    {
                        modulo = (dividend - 1) % 26;
                        columnLetter = Convert.ToChar(65 + modulo) + columnLetter;
                        dividend = (int)((dividend - modulo) / 26);
                    }

                    return columnLetter;
                }
                
                Worksheet labourVar1SalesSheet = labourVar1Workbook.Worksheets["Sales"];

                int salesColumn = labourVar1SalesSheet.UsedRange.Columns.Count - 1;

                string salesColumnLetter = GetColumnLetter(salesColumn);

                Excel.Range salesAddColumn = labourVar1SalesSheet.Columns[salesColumn];
                salesAddColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                labourVar1SalesSheet.Cells[3,salesColumn].Value = date;
                labourVar1SalesSheet.Cells[4, salesColumn].Formula = $"=SUM({salesColumnLetter}5:{salesColumnLetter}60)";
                labourVar1SalesSheet.Range[$"{salesColumnLetter}5:{salesColumnLetter}{labourVar1SalesSheet.UsedRange.Rows.Count}"].Formula = $"=IFERROR((VLOOKUP($B5,'Week {previousWeekNbr}'!$A:$C,3,FALSE)+VLOOKUP($B5,'Week {currentWeekNbr}'!$A:$C,3,FALSE)),0)";

                Worksheet ffcglSheet = ffcglWorkbook.Worksheets["FFCGL"];
                Worksheet labourVar1FfcglSheet = labourVar1Workbook.Worksheets["FFCGL"];

                List<FfcglData> ffcglList = new List<FfcglData> ();
                int ffcglLastRow = ffcglSheet.Cells[ffcglSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;
                int labourVar1FfcglSheetLastRow1 = labourVar1FfcglSheet.Cells[labourVar1FfcglSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 1;
                int labourVar1FfcglSheetLastRow2 = labourVar1FfcglSheetLastRow1;

                for (int i = 1; i <= ffcglLastRow; i++)
                {
                    string entity = Convert.ToString(ffcglSheet.Cells[i, 1].Value);
                    string store = Convert.ToString(ffcglSheet.Cells[i, 2].Value);
                    string jobDesp = Convert.ToString(ffcglSheet.Cells[i, 3].Value);
                    string desc = Convert.ToString(ffcglSheet.Cells[i, 4].Value);
                    string amount = Convert.ToString(ffcglSheet.Cells[i, 5].Value);

                    if (!string.IsNullOrWhiteSpace(entity) && (entity.Contains("DFG") || entity.Contains("FSH") || entity.Contains("SUN")))
                    {
                        if (!string.IsNullOrWhiteSpace(desc) && desc.StartsWith("E"))
                        {
                            FfcglData rowData = new FfcglData
                            {
                                Entity = entity,
                                Store = store,
                                JobDesp = jobDesp,
                                Desc = desc,
                                Amount = amount,
                                
                            };

                            ffcglList.Add(rowData);
                        }
                    }
                }

                foreach (var data in ffcglList)
                {
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 2].Value = data.Entity;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 3].Value = data.Store;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 4].Value = data.JobDesp;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 5].Value = data.Desc;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 6].Value = data.Amount;
                    labourVar1FfcglSheetLastRow1++;
                }

                labourVar1FfcglSheet.Range[$"A{labourVar1FfcglSheetLastRow2}:A{labourVar1FfcglSheetLastRow1 - 1}"].Value = date;

                Worksheet pivotSheet = labourVar1Workbook.Worksheets["Data"];
                PivotTable pivotTable = pivotSheet.PivotTables(1);
                pivotTable.RefreshTable();

                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                string month = parsedDate.ToString("MMM", CultureInfo.InvariantCulture);
                string day = parsedDate.Day.ToString();

                var pivotDate = $"{day}-{month}";

                Excel.Range pivotRange = pivotSheet.PivotTables(1).TableRange1;

                string pivotColumnLetter = string.Empty;

                foreach (Excel.Range cell in pivotRange.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString() == pivotDate)
                    {
                        //Console.WriteLine($"Found date '{pivotDate}' at cell: {cell.Address}");
                        string cellAddress = cell.Address.ToString();
                        pivotColumnLetter = new string(cellAddress.Where(char.IsLetter).ToArray());
                        break; 
                    }
                }

                Worksheet labourVar1LaborsSheet = labourVar1Workbook.Worksheets["Labors"];

                int laborsColumn = labourVar1LaborsSheet.UsedRange.Columns.Count - 2;

                string laborsColumnLetter = GetColumnLetter(laborsColumn);

                Excel.Range laborsAddColumn = labourVar1LaborsSheet.Columns[laborsColumn];
                laborsAddColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                labourVar1LaborsSheet.Cells[3, laborsColumn].Value = date;
                labourVar1LaborsSheet.Cells[4, laborsColumn].Formula = $"=SUM({laborsColumnLetter}5:{laborsColumnLetter}60)";
                labourVar1LaborsSheet.Range[$"{laborsColumnLetter}5:{laborsColumnLetter}60"].Formula = $"=XLOOKUP($B5,Data!$A:$A,Data!{pivotColumnLetter}:{pivotColumnLetter},0,0,1)";

                //labourVar1Workbook.SaveAs(@"C: \Users\Nimap\Downloads\Labor var Van - Copy\test\Labor Var 2023 - 11 - 20.xlsx");
                //labourVar1Workbook.Save();
                
                // Second file - Cj labour

                string cjMonth = parsedDate.ToString("MM", CultureInfo.InvariantCulture);
                string cjDay = parsedDate.ToString("dd", CultureInfo.InvariantCulture);

                var cjDate = $"{cjMonth}{cjDay}";

                DateTime twoWeeksPreviousDate = parsedDate.AddDays(-14);

                string previousCjMonth = twoWeeksPreviousDate.ToString("MM", CultureInfo.InvariantCulture);
                string previousCjday = twoWeeksPreviousDate.ToString("dd", CultureInfo.InvariantCulture);

                var cjTwoWeeksPreviousDate = $"{previousCjMonth}{previousCjday}";

                Excel.Worksheet previousSheet = null;
                foreach (Excel.Worksheet sheet in cjLabourWorkbook.Worksheets)
                {
                    if (sheet.Name == $"Labor Standard Var {cjTwoWeeksPreviousDate}")
                    {
                        previousSheet = sheet;
                        break;
                       
                    }
                }

                previousSheet.Copy(After: previousSheet);
                Excel.Worksheet newSheet = (Excel.Worksheet)cjLabourWorkbook.Sheets[previousSheet.Index + 1];
                newSheet.Name = $"Labor Standard Var {cjDate}";

                Excel.Range cellC4 = newSheet.Range["C4"];
                cellC4.Value = date;

                Worksheet cjLabourSalesSheet = cjLabourWorkbook.Worksheets["Sales"];
                var cjPreviousLabourDate = $"PPE {previousCjMonth}/{previousCjday}";

                Excel.Range row2 = cjLabourSalesSheet.Rows[2];
                int cjLabourColumn = -1;
                foreach (Excel.Range cell in row2.Cells)
                {
                    if (cell.Value == cjPreviousLabourDate)
                    {
                        cjLabourColumn = cell.Column; 
                        break;
                    }
                }

                string cjLaborsColumnLetter = GetColumnLetter(cjLabourColumn + 1);

                Excel.Range cjLaborsAddColumn = cjLabourSalesSheet.Columns[cjLabourColumn + 1];
                cjLaborsAddColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                cjLabourSalesSheet.Cells[2, cjLabourColumn + 1].Value = $"PPE {cjMonth}/{cjDay}";
                cjLabourSalesSheet.Cells[3, cjLabourColumn + 1].Value = "$";

                Excel.Range sourceLabourRange = labourVar1LaborsSheet.Range[$"{laborsColumnLetter}5:{laborsColumnLetter}{labourVar1LaborsSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                Excel.Range targetLabourRange = cjLabourSalesSheet.Range[$"{cjLaborsColumnLetter}5:{cjLaborsColumnLetter}{cjLabourSalesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                sourceLabourRange.Copy(Type.Missing);
                targetLabourRange.PasteSpecial(XlPasteType.xlPasteValues);

                cjLabourSalesSheet.Cells[4, cjLabourColumn + 1].Formula = $"=SUM({cjLaborsColumnLetter}5:{cjLaborsColumnLetter}60)";

                var cjPreviousSalesDate = $"Ending {previousCjMonth}/{previousCjday}";

                Excel.Range row3 = cjLabourSalesSheet.Rows[3];
                int cjSalesColumn = -1;
                foreach (Excel.Range cell in row3.Cells)
                {
                    if (cell.Value == cjPreviousSalesDate)
                    {
                        cjSalesColumn = cell.Column;
                        break;
                    }
                }

                string cjSalesColumnLetter = GetColumnLetter(cjSalesColumn + 1);

                Excel.Range cjSalesAddColumn = cjLabourSalesSheet.Columns[cjSalesColumn + 1];
                cjSalesAddColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                cjLabourSalesSheet.Cells[3, cjSalesColumn + 1].Value = $"Ending {cjMonth}/{cjDay}";

                Excel.Range sourceSalesRange = labourVar1SalesSheet.Range[$"{salesColumnLetter}5:{salesColumnLetter}{labourVar1SalesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                Excel.Range targetSalesRange = cjLabourSalesSheet.Range[$"{cjSalesColumnLetter}5:{cjSalesColumnLetter}{cjLabourSalesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                sourceSalesRange.Copy(Type.Missing);
                targetSalesRange.PasteSpecial(XlPasteType.xlPasteValues);

                cjLabourSalesSheet.Cells[4, cjSalesColumn + 1].Formula = $"=SUM({cjSalesColumnLetter}5:{cjSalesColumnLetter}60)";

                newSheet.Range[$"C9:C62"].Formula = $"=INDEX(Sales!{cjSalesColumnLetter}:{cjSalesColumnLetter},MATCH('Labor Standard Var 1204'!B9,Sales!B:B,0))";
                newSheet.Range[$"F9:F62"].Formula = $"=INDEX(Sales!{cjLaborsColumnLetter}:{cjLaborsColumnLetter},MATCH('Labor Standard Var 1204'!B9,Sales!B:B,0))";

                // third file - labour var van 2

                Worksheet summarySheet = labourVar2Workbook.Worksheets["Summary"];

                List<CjLabourStandardData> cjLabourStandardDatas = new List<CjLabourStandardData>();

                for (int i = 9; i < 63; i++)
                {
                    string cjStore = Convert.ToString(newSheet.Cells[i, 2].Value);

                    string salesCellValue = Convert.ToString(newSheet.Cells[i, 3].Value);
                    salesCellValue = salesCellValue.Replace("$", ""); 
                    double salesValue = 0.0;
                    if (double.TryParse(salesCellValue, out double parsedSalesValue))
                    {
                        salesValue = parsedSalesValue / 1000; 
                    }

                    string labourCellValue = Convert.ToString(newSheet.Cells[i, 5].Value);
                    labourCellValue = labourCellValue.Replace("%", ""); 
                    double labourDouble = 0.0;
                    if (double.TryParse(labourCellValue, out double parsedLabourValue))
                    {
                        labourDouble = parsedLabourValue * 100;
                        labourCellValue = labourDouble.ToString(); 
                    }

                    string actualCellValue = Convert.ToString(newSheet.Cells[i, 7].Value);
                    actualCellValue = actualCellValue.Replace("%", ""); 
                    double actualDouble = 0.0;
                    if (double.TryParse(actualCellValue, out double parsedActualValue))
                    {
                        actualDouble = parsedActualValue * 100;
                        actualCellValue = actualDouble.ToString(); 
                    }

                    string varCellValue = Convert.ToString(newSheet.Cells[i, 14].Value);
                    varCellValue = varCellValue.Replace("%", ""); 
                    double varDouble = 0.0;
                    if (double.TryParse(varCellValue, out double parsedVarValue))
                    {
                        varDouble = parsedVarValue * 100;
                        varCellValue = varDouble.ToString(); 
                    }

                    CjLabourStandardData data = new CjLabourStandardData
                    {
                        CjStore = cjStore,
                        Sales = salesValue,
                        Labour = labourCellValue,
                        Actual = actualCellValue,
                        Var = varCellValue
                    };

                    cjLabourStandardDatas.Add(data);
                }
                
                var summaryRow = 5;
                foreach (var data in cjLabourStandardDatas)
                {
                    summarySheet.Cells[summaryRow, 2].Value = data.CjStore;
                    summarySheet.Cells[summaryRow, 3].Value = data.Sales;
                    summarySheet.Cells[summaryRow, 4].Value = data.Labour;
                    summarySheet.Cells[summaryRow, 5].Value = data.Actual;
                    summarySheet.Cells[summaryRow, 6].Value = data.Var;
                    summaryRow++;
                }

                Excel.Range columnsToInsert = summarySheet.Columns["T:T"].Resize[Missing.Value, 3]; 
                columnsToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove); 

                Excel.Range sourceRange = summarySheet.Range["Q:Q,R:R"]; 
                Excel.Range destinationRange = summarySheet.Range["T:U"]; 

                sourceRange.Copy(destinationRange);

                Excel.Range cellT4 = summarySheet.Cells[4, 20]; 
                Excel.Range cellU4 = summarySheet.Cells[4, 21]; 
                Excel.Range mergedRange = summarySheet.Range[cellT4, cellU4];
                mergedRange.Merge();

                Excel.Range cellQ4 = summarySheet.Cells[4, 17]; 
                Excel.Range cellR4 = summarySheet.Cells[4, 18]; 
                cellQ4.MergeArea.Value = date;

                Worksheet cjListing = labourVar2Workbook.Worksheets[1];
                Worksheet siteList = storeList.Worksheets[1];

                Excel.Range clearRange = cjListing.Range["A1:N" + cjListing.Rows.Count];
                clearRange.Clear();

                Excel.Range copyRange = siteList.Range["A1:N" + siteList.Rows.Count];
                copyRange.Copy(clearRange);

                











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
