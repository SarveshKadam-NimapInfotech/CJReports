using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using static OfficeOpenXml.ExcelErrorValue;
using Excel = Microsoft.Office.Interop.Excel;

namespace LabourVarVanChanges
{
    public static class HelperMethod
    {
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
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

            string labourVar1FilePath = @"C:\Users\Nimap\Documents\Labour Van\input files\Labor Var 2024-01-15.xlsx";

            string labourVar2FilePath = @"C:\Users\Nimap\Documents\Labour Van\input files\Labor Var 2024.01.15.xlsx";

            string ffcglFilePath = @"C:\Users\Nimap\Documents\Labour Van\input files\FFCGL0129.xlsx";

            string cjLabourFilePath = @"C:\Users\Nimap\Documents\Labour Van\input files\CJ Labor Standard 2024-01-15.xlsx";

            string labourVarTrendFilePath = @"C:\Users\Nimap\Documents\Labour Van\input files\Labor Var Trend since 8.29.22.xlsx";

            string weeklySalesFilePath = @"C:\Users\Nimap\Documents\Labour Van\input files\Week 4 Sales 2024-01-29.xlsm";

            string storeListFilePath = @"C:\Users\Public\Documents\StoreList.xlsx";
            Excel.Workbook storeList = excelApp.Workbooks.Open(storeListFilePath);

            Excel.Workbook labourVar1Workbook = excelApp.Workbooks.Open(labourVar1FilePath, CorruptLoad: XlCorruptLoad.xlExtractData);

            Excel.Workbook labourVar2Workbook = excelApp.Workbooks.Open(labourVar2FilePath);

            Excel.Workbook ffcglWorkbook = excelApp.Workbooks.Open(ffcglFilePath);

            Excel.Workbook cjLabourWorkbook = excelApp.Workbooks.Open(cjLabourFilePath, CorruptLoad: XlCorruptLoad.xlExtractData);

            Excel.Workbook labourVarTrendWorkbook = excelApp.Workbooks.Open(labourVarTrendFilePath);

            Excel.Workbook weeklySalesWorkbook = excelApp.Workbooks.Open(weeklySalesFilePath);

            try
            {
                var date = "01/29/2024";

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

                Worksheet labourVar1SalesSheet = labourVar1Workbook.Worksheets["Sales"];

                int salesColumn = labourVar1SalesSheet.UsedRange.Columns.Count - 1;

                string salesColumnLetter = HelperMethod.GetExcelColumnName(salesColumn);

                Excel.Range salesAddColumn = labourVar1SalesSheet.Columns[salesColumn];
                salesAddColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                labourVar1SalesSheet.Cells[3, salesColumn].Value = date;
                labourVar1SalesSheet.Cells[4, salesColumn].Formula = $"=SUM({salesColumnLetter}5:{salesColumnLetter}60)";
                labourVar1SalesSheet.Range[$"{salesColumnLetter}5:{salesColumnLetter}{labourVar1SalesSheet.UsedRange.Rows.Count}"].Formula = $"=IFERROR((VLOOKUP($B5,'Week {previousWeekNbr}'!$A:$C,3,FALSE)+VLOOKUP($B5,'Week {currentWeekNbr}'!$A:$C,3,FALSE)),0)";

                labourVar1SalesSheet.Columns.AutoFit();

                Worksheet ffcglSheet = ffcglWorkbook.Worksheets[1];
                Worksheet labourVar1FfcglSheet = labourVar1Workbook.Worksheets["FFCGL"];

                List<FfcglData> ffcglList = new List<FfcglData>();
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
                        if (!string.IsNullOrWhiteSpace(desc) && desc.StartsWith("E") && !desc.Contains("Bonus") && !desc.Contains("Vacation"))
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

                //int dateRow = 2;
                //while (dateRow < labourVar1FfcglSheetLastRow1)
                //{
                //    string dateOfRow = Convert.ToString(labourVar1FfcglSheet.Cells[dateRow, 1].Value);
                //    if (dateOfRow.Contains("12/19/2022"))
                //    {
                //        Excel.Range rowToDelete = labourVar1FfcglSheet.Rows[dateRow];
                //        rowToDelete.Delete();
                //    }
                //    if (dateOfRow.Contains("1/2/2023"))
                //    {
                //        break;
                //    }

                //}



                Worksheet pivotSheet = labourVar1Workbook.Worksheets["Data"];

                //PivotTable pivotTable = pivotSheet.PivotTables(1);
                //pivotTable.RefreshTable();

                //PivotTables pivotTables = pivotSheet.PivotTables();
                //int pivotTablesCount = pivotTables.Count;
                //if (pivotTablesCount > 0)
                //{
                //    for (int i = 1; i <= pivotTablesCount; i++)
                //    {
                //        pivotTables.Item(i).RefreshTable(); 
                //    }
                //}
                //CustomLogging._logger.Info("pivot refreshed");


                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                string month = parsedDate.ToString("MMM", CultureInfo.InvariantCulture);
                string day = parsedDate.Day.ToString();
                int year = parsedDate.Year;

                //var pivotDate = $"{day}-{month}";

                pivotSheet.Cells.Clear();

                Worksheet Pivot = pivotSheet;
                Range pivotData = labourVar1FfcglSheet.Range[$"A:F"];

                PivotTable pivotTable = labourVar1FfcglSheet.PivotTableWizard(XlPivotTableSourceType.xlDatabase, pivotData, Pivot.Range["A2"], "PIVOT");

                PivotField amountFields = pivotTable.PivotFields("Amount");
                amountFields.Orientation = XlPivotFieldOrientation.xlDataField;
                amountFields.Function = XlConsolidationFunction.xlSum;

                PivotField descFields = pivotTable.PivotFields("PPE");
                descFields.Orientation = XlPivotFieldOrientation.xlColumnField;

                //descFields.PivotFilters.Add2(
                //Type: Excel.XlPivotFilterType.xlCaptionDoesNotEqual,
                //Value1: "(blank)"
                //);

                //descFields.NumberFormat = "d-mmm";
                PivotField storeFields = pivotTable.PivotFields("Store");
                storeFields.Orientation = XlPivotFieldOrientation.xlRowField;


                pivotTable.TableStyle2 = "PivotStyleLight16";

                var pivotColumnRange = pivotSheet.UsedRange.Columns.Count;
                int column = 1;
                while (column <= pivotColumnRange)
                {
                    pivotSheet.Cells[1, column].Value = column;
                    column++;
                }

                string pivotColumnLetter = HelperMethod.GetExcelColumnName(pivotColumnRange - 2);


                //Excel.Range pivotRange = pivotSheet.PivotTables(1).TableRange1;

                //string pivotColumnLetter = string.Empty;

                //foreach (Excel.Range cell in pivotRange.Cells)
                //{
                //    if (cell.Value != null && cell.Value.ToString() == pivotDate)
                //    {
                //        string cellAddress = cell.Address.ToString();
                //        pivotColumnLetter = new string(cellAddress.Where(char.IsLetter).ToArray());
                //        break;
                //    }
                //}

                Worksheet labourVar1LaborsSheet = labourVar1Workbook.Worksheets["Labors"];

                int laborsColumn = labourVar1LaborsSheet.UsedRange.Columns.Count - 2;

                string laborsColumnLetter = HelperMethod.GetExcelColumnName(laborsColumn);

                Excel.Range laborsAddColumn = labourVar1LaborsSheet.Columns[laborsColumn];
                laborsAddColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                labourVar1LaborsSheet.Cells[3, laborsColumn].Value = date;
                labourVar1LaborsSheet.Cells[4, laborsColumn].Formula = $"=SUM({laborsColumnLetter}5:{laborsColumnLetter}60)";
                labourVar1LaborsSheet.Range[$"{laborsColumnLetter}5:{laborsColumnLetter}60"].Formula = $"=IFERROR(INDEX(Data!{pivotColumnLetter}:{pivotColumnLetter},MATCH($B5,Data!$A:$A,0)),0)";

                labourVar1LaborsSheet.Columns.AutoFit();


                // Second file - Cj labour

                string cjMonth = parsedDate.ToString("MM", CultureInfo.InvariantCulture);
                string cjDay = parsedDate.ToString("dd", CultureInfo.InvariantCulture);
                string cjYear = parsedDate.ToString("yyyy", CultureInfo.InvariantCulture);

                var cjDate = $"{cjMonth}{cjDay}";

                DateTime twoWeeksPreviousDate = parsedDate.AddDays(-14);

                string previousCjMonth = twoWeeksPreviousDate.ToString("MM", CultureInfo.InvariantCulture);
                string previousCjday = twoWeeksPreviousDate.ToString("dd", CultureInfo.InvariantCulture);
                string previousCjYear = twoWeeksPreviousDate.ToString("yyyy", CultureInfo.InvariantCulture);

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

                string cjLaborsColumnLetter = HelperMethod.GetExcelColumnName(cjLabourColumn + 1);

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

                string cjSalesColumnLetter = HelperMethod.GetExcelColumnName(cjSalesColumn + 1);

                Excel.Range cjSalesAddColumn = cjLabourSalesSheet.Columns[cjSalesColumn + 1];
                cjSalesAddColumn.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                cjLabourSalesSheet.Cells[3, cjSalesColumn + 1].Value = $"Ending {cjMonth}/{cjDay}";

                Excel.Range sourceSalesRange = labourVar1SalesSheet.Range[$"{salesColumnLetter}5:{salesColumnLetter}{labourVar1SalesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                Excel.Range targetSalesRange = cjLabourSalesSheet.Range[$"{cjSalesColumnLetter}5:{cjSalesColumnLetter}{cjLabourSalesSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                sourceSalesRange.Copy(Type.Missing);
                targetSalesRange.PasteSpecial(XlPasteType.xlPasteValues);

                cjLabourSalesSheet.Cells[4, cjSalesColumn + 1].Formula = $"=SUM({cjSalesColumnLetter}5:{cjSalesColumnLetter}60)";

                cjLabourSalesSheet.Columns.AutoFit();

                newSheet.Range[$"C9:C62"].Formula = $"=INDEX(Sales!{cjSalesColumnLetter}:{cjSalesColumnLetter},MATCH('Labor Standard Var 1204'!B9,Sales!B:B,0))";
                newSheet.Range[$"F9:F62"].Formula = $"=INDEX(Sales!{cjLaborsColumnLetter}:{cjLaborsColumnLetter},MATCH('Labor Standard Var 1204'!B9,Sales!B:B,0))";

                newSheet.Columns.AutoFit();

                Worksheet cjLabourStandardSheet = cjLabourWorkbook.Worksheets[$"Labor Standard Var {cjDate}"];

                Worksheet summarySheet = labourVar2Workbook.Worksheets["Summary"];

                Excel.Range columnsToInsert = summarySheet.Columns["T:T"].Resize[Missing.Value, 3];
                columnsToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                Excel.Range cellQ4 = summarySheet.Cells[4, 17];
                Excel.Range cellR4 = summarySheet.Cells[4, 18];
                Excel.Range unmergedRange = summarySheet.Range[cellQ4, cellR4];
                unmergedRange.UnMerge();
                Excel.Range sourceRange = summarySheet.Range["Q:R"];
                Excel.Range destinationRange = summarySheet.Range["T:U"];

                sourceRange.Copy(Type.Missing);
                destinationRange.PasteSpecial(XlPasteType.xlPasteFormats);

                sourceRange.Copy(Type.Missing);
                destinationRange.PasteSpecial(XlPasteType.xlPasteValues);

                int cjLabourStandardLastRow = 0;

                for (int rows = 9; rows < cjLabourStandardSheet.Rows.Count; rows++)
                {
                    if (cjLabourStandardSheet.Cells[rows, 2].Value == null)
                    {
                        cjLabourStandardLastRow = rows;
                        break;

                    }
                }

                List<CjLabourStandardData> cjLabourStandardDatas = new List<CjLabourStandardData>();

                for (int i = 9; i < cjLabourStandardLastRow; i++)
                {
                    string cjStore = Convert.ToString(cjLabourStandardSheet.Cells[i, 2].Value);

                    string salesCellValue = Convert.ToString(cjLabourStandardSheet.Cells[i, 3].Value);
                    salesCellValue = salesCellValue.Replace("$", "");
                    double salesValue = 0.0;
                    if (double.TryParse(salesCellValue, out double parsedSalesValue))
                    {
                        salesValue = parsedSalesValue / 1000;
                    }

                    string actualCellValue = Convert.ToString(cjLabourStandardSheet.Cells[i, 7].Value);
                    actualCellValue = actualCellValue.Replace("%", "");
                    double actualDouble = 0.0;
                    if (double.TryParse(actualCellValue, out double parsedActualValue))
                    {
                        actualDouble = parsedActualValue * 100;
                        actualCellValue = actualDouble.ToString();
                    }

                    CjLabourStandardData data = new CjLabourStandardData
                    {
                        CjStore = cjStore,
                        Sales = salesValue,
                        Actual = actualCellValue,
                    };

                    cjLabourStandardDatas.Add(data);
                }

                var summaryRow = 5;
                foreach (var data in cjLabourStandardDatas)
                {
                    summarySheet.Cells[summaryRow, 2].Value = data.CjStore;
                    summarySheet.Cells[summaryRow, 3].Value = data.Sales;
                    summarySheet.Cells[summaryRow, 5].Value = data.Actual;
                    summaryRow++;
                }

                Worksheet changingListSheet = cjLabourWorkbook.Worksheets["Changing list"];

                List<string> regList = new List<string>();

                for (int i = 9; i < cjLabourStandardLastRow; i++)
                {
                    regList.Add(Convert.ToString(changingListSheet.Cells[i, 13].Value));

                }

                var changeListRow = 9;
                foreach (var data in regList)
                {
                    cjLabourStandardSheet.Cells[changeListRow, 13].Value = data;
                    changeListRow++;
                }


                //third file -labour var van 2

                //var targetPath = new TargetPathForLaborVar
                //{
                //    labourVar2Workbook = Path.Combine(outputPath, $"Labor Var {cjYear}.{cjMonth}.{cjDay}_{Guid.NewGuid()}.xlsx"),
                //    //$"{baseDirectory}\\Labor Var {cjYear}.{cjMonth}.{cjDay}.xlsx",
                //    labourVar1Workbook = Path.Combine(outputPath, $"Labor Var {cjYear}-{cjMonth}-{cjDay}_{Guid.NewGuid()}.xlsx"),
                //    cjLabourWorkbook = Path.Combine(outputPath, $"CJ Labor Standard {cjYear}-{cjMonth}-{cjDay}_{Guid.NewGuid()}.xlsx"),
                //    labourVarTrendWorkbook = Path.Combine(outputPath, $"Labor Var Trend since 8.29.22_{Guid.NewGuid()}.xlsx"),
                //    //ffcglWorkbook = "C:\\Users\\Public\\Documents\\FFCGL.xlsx",

                //};

                //cjLabourWorkbook.SaveAs(cjLabourWorkbook);
                //cjLabourWorkbook.Close();
                //Marshal.ReleaseComObject(cjLabourWorkbook);


                //List<double> firstValue = firstValue = new List<double>();
                //List<double> secondValue = secondValue = new List<double>();

                //using (var excelPackage = new ExcelPackage(new FileInfo(cjLabourFilePath)))
                //{

                //    var worksheet = excelPackage.Workbook.Worksheets[$"Labor Standard Var {cjDate}"];

                //    int lastRowUsed = 0;
                //    for (int rows = 9; rows < worksheet.Dimension.End.Row; rows++)
                //    {
                //        if (worksheet.Cells[rows, 5].Value == null)
                //        {
                //            lastRowUsed = rows;
                //            break;
                //        }

                //    }

                //    ExcelRangeBase range1 = worksheet.Cells[$"E9:E{lastRowUsed - 1}"];
                //    ExcelRangeBase range2 = worksheet.Cells[$"N9:N{lastRowUsed - 1}"];


                //    range1.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
                //    range2.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });


                //    for (int rows = 9; rows <= lastRowUsed; rows++)
                //    {
                //        if (worksheet.Cells[rows, 5].Value != null && worksheet.Cells[rows, 5].Value is double)
                //        {
                //            firstValue.Add(Convert.ToDouble(worksheet.Cells[rows, 5].Value));
                //        }
                //        if (worksheet.Cells[rows, 14].Value != null && worksheet.Cells[rows, 14].Value is double)
                //        {
                //            secondValue.Add(Convert.ToDouble(worksheet.Cells[rows, 14].Value));
                //        }
                //    }

                //}

                //summaryRow = 5;
                //foreach (var data in firstValue)
                //{

                //    double roundedValue = Math.Round(data * 100, 1);
                //    summarySheet.Cells[summaryRow, 4].Value = roundedValue;
                //    summaryRow++;
                //}

                //summaryRow = 5;
                //foreach (var data in secondValue)
                //{

                //    double roundedValue = Math.Round(data * 100, 1);
                //    summarySheet.Cells[summaryRow, 6].Value = roundedValue;
                //    summaryRow++;
                //}

                //Excel.Range columnsToInsert = summarySheet.Columns["T:T"].Resize[Missing.Value, 3];
                //columnsToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);

                //Excel.Range cellQ4 = summarySheet.Cells[4, 17];
                //Excel.Range cellR4 = summarySheet.Cells[4, 18];
                //Excel.Range unmergedRange = summarySheet.Range[cellQ4, cellR4];
                //unmergedRange.UnMerge();
                //Excel.Range sourceRange = summarySheet.Range["Q:R"];
                //Excel.Range destinationRange = summarySheet.Range["T:U"];

                //sourceRange.Copy(Type.Missing);
                //destinationRange.PasteSpecial(XlPasteType.xlPasteFormats);

                //sourceRange.Copy(Type.Missing);
                //destinationRange.PasteSpecial(XlPasteType.xlPasteValues);


                Excel.Range cellT4 = summarySheet.Cells[4, 20];
                Excel.Range cellU4 = summarySheet.Cells[4, 21];
                Excel.Range mergedRange = summarySheet.Range[cellT4, cellU4];
                mergedRange.Merge();
                cellT4.MergeArea.Value = $"{previousCjMonth}/{previousCjday}/{previousCjYear}";

                Range sortURange = summarySheet.Range["T5:U15"];
                sortURange.Sort(sortURange.Columns[2], XlSortOrder.xlAscending, Type.Missing, Type.Missing);

                unmergedRange.Merge();
                cellQ4.MergeArea.Value = date;


                Worksheet cjListing = labourVar2Workbook.Worksheets[1];
                Worksheet siteList = storeList.Worksheets[1];

                Excel.Range copyRange = siteList.Range["A1:N" + siteList.Rows.Count];
                copyRange.Copy(Type.Missing);

                Excel.Range pasteRange = cjListing.Range["A1:N" + cjListing.Rows.Count];
                pasteRange.PasteSpecial(XlPasteType.xlPasteAll);

                Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

                int row = 6;
                for (int i = 1; i <= cjListing.UsedRange.Rows.Count; i++)
                {
                    var regionValue = cjListing.Cells[i, 1].Value;
                    if ((regionValue == null) || !(regionValue is string))
                    {
                        continue;
                    }
                    string region = Convert.ToString(cjListing.Cells[i, 1].Value);
                    if (region.ToLower().Trim().Contains("region"))
                    {
                        row = i;
                        break;
                    }
                }

                while (cjListing.Cells[row, 1].Value != null)
                {
                    string cellValue = Convert.ToString(cjListing.Cells[row, 1].Value);
                    if (cellValue == "North")
                    {
                        break;
                    }

                    if (cellValue.StartsWith("Dist") || cellValue.StartsWith("D"))
                    {
                        string key = cellValue;

                        if (cellValue.StartsWith("D") && cellValue.Length > 1 && !char.IsLetter(cellValue[1]))
                        {
                            key = "Dist " + cellValue.Substring(1);
                        }

                        List<string> values = new List<string>();

                        while (cjListing.Cells[++row, 1].Value != null)
                        {
                            string nextCellValue = Convert.ToString(cjListing.Cells[row, 1].Value);

                            if (nextCellValue.StartsWith("Dist") || nextCellValue.StartsWith("D") || nextCellValue.StartsWith("Region") || nextCellValue == "North")
                            {
                                break;
                            }

                            values.Add(nextCellValue);
                        }

                        dict.Add(key, values);
                    }
                    else
                    {
                        row++;
                    }
                }

                for (int i = 1; i <= 11; i++)
                {
                    string sheetNumber = i.ToString().PadLeft(2, '0');

                    foreach (Excel.Worksheet sheet in labourVar2Workbook.Worksheets)
                    {
                        if (sheet.Name.Contains(sheetNumber))
                        {
                            int distRow = 4;
                            while (sheet.Cells[distRow, 2].Value != null)
                            {
                                sheet.Cells[distRow, 2].Value = null;
                                distRow++;
                            }

                            List<string> distValues = dict.ContainsKey($"Dist {i}") ? dict[$"Dist {i}"] : new List<string>();
                            distRow = 4;
                            foreach (string value in distValues)
                            {
                                if (sheet.Cells[distRow, 1].Value == null)
                                {
                                    sheet.Rows[distRow - 1].Copy(Type.Missing);
                                    sheet.Rows[distRow - 1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                                    sheet.Rows[distRow].PasteSpecial(XlPasteType.xlPasteAll);
                                }

                                sheet.Cells[distRow, 2].Value = value;

                                distRow++;
                            }

                            while (sheet.Cells[distRow, 1].Value != null)
                            {
                                sheet.Rows[distRow].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            }
                        }
                    }
                }

                Range sortRange = summarySheet.Range[$"B{5}:F{summarySheet.UsedRange.Rows.Count}"];
                sortRange.Sort(sortRange.Columns[5], XlSortOrder.xlAscending);


                foreach (Excel.Worksheet sheet in labourVar2Workbook.Worksheets)
                {
                    if (sheet.Name.Contains("District"))
                    {

                        Range sortDistrictRange = sheet.Range[$"A{4}:F{sheet.UsedRange.Rows.Count}"];
                        sortDistrictRange.Sort(sortDistrictRange.Columns[6], XlSortOrder.xlAscending);

                        sheet.Cells[2, 1].MergeArea.Value = date;


                    }
                }

                // Adding the Dynamic code for district in the summary sheet

                List<string> districts = new List<string>();

                for (int i = 1; i <= cjListing.UsedRange.Rows.Count; i++)
                {
                    string cellValue = Convert.ToString(cjListing.Cells[i, 1].Value);
                    if (cellValue == "North")
                    {
                        break;
                    }

                    if (cellValue.StartsWith("Dist"))
                    {
                        string districtNumber = cellValue.Substring(4).Trim();

                        if (int.TryParse(districtNumber, out int _))
                        {
                            districts.Add("D" + districtNumber);
                        }

                    }
                    if (cellValue.StartsWith("D") && !cellValue.StartsWith("Dist"))
                    {
                        districts.Add(cellValue);
                    }

                }

                Range clearDRange = summarySheet.Range["Q5:R15"];
                clearDRange.ClearContents();

                var districtStartRow = 5;
                foreach (string district in districts)
                {
                    summarySheet.Cells[districtStartRow, 17].Value = district;
                   
                    string districtNumber = district.Substring(1).Trim();

                    if (int.TryParse(districtNumber, out int _))
                    {
                        summarySheet.Cells[districtStartRow, 17].Formula = $"='District {districtNumber}'!F4";

                    }

                    districtStartRow++;

                }



                Range sortDRange = summarySheet.Range["Q5:R15"];
                sortDRange.Sort(sortDRange.Columns[2], XlSortOrder.xlAscending, Type.Missing, Type.Missing);

                summarySheet.Cells[1, 1].MergeArea.Value = $"{cjMonth}-{cjDay}-{cjYear} CARLS JR LABOR VARIANCE";

                summarySheet.Cells[1, 8].MergeArea.Value = $"{cjMonth}-{cjDay}-{cjYear} CARLS JR LABOR VARIANCE";


                // Fourth File - Labour var trends

                Worksheet labourVarTrendSheet1 = labourVarTrendWorkbook.Worksheets["Data"];
                Worksheet labourVarTrendSheet2 = labourVarTrendWorkbook.Worksheets["BI-Weekly Trend"];

                var labourVarTrendLastRow = labourVarTrendSheet1.UsedRange.Rows.Count + 2;

                Excel.Range labourVarCopyRange1 = summarySheet.Range["A1:O4"];
                Excel.Range labourVarPasteRange1 = labourVarTrendSheet1.Range[$"B{labourVarTrendLastRow} : P{labourVarTrendLastRow + 3}"];

                labourVarCopyRange1.Copy(Type.Missing);
                labourVarPasteRange1.PasteSpecial(Excel.XlPasteType.xlPasteFormats);

                labourVarTrendSheet1.Cells[labourVarTrendLastRow, 2].MergeArea.Value = $"{cjMonth}-{cjDay}-{cjYear} CARLS JR LABOR VARIANCE";
                labourVarTrendSheet1.Cells[labourVarTrendLastRow, 9].MergeArea.Value = $"{cjMonth}-{cjDay}-{cjYear} CARLS JR LABOR VARIANCE";
                labourVarTrendSheet1.Cells[labourVarTrendLastRow + 1, 2].MergeArea.Value = "'Average";
                labourVarTrendSheet1.Cells[labourVarTrendLastRow + 1, 9].MergeArea.Value = "'Weighted Average";
                labourVarTrendSheet1.Cells[labourVarTrendLastRow + 3, 2].MergeArea.Value = "Comp. Avg.";
                labourVarTrendSheet1.Cells[labourVarTrendLastRow + 3, 9].MergeArea.Value = "Comp. Avg.";
                Excel.Range labourVarRowCopy1 = summarySheet.Range["A3:O3"];
                Excel.Range labourVarRowPaste1 = labourVarTrendSheet1.Range[$"B{labourVarTrendLastRow + 2} : P{labourVarTrendLastRow + 2}"];
                labourVarRowCopy1.Copy(Type.Missing);
                labourVarRowPaste1.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                Excel.Range labourVarRowCopy2 = summarySheet.Range["C4:F4"];
                Excel.Range labourVarRowPaste2 = labourVarTrendSheet1.Range[$"D{labourVarTrendLastRow + 3} : G{labourVarTrendLastRow + 3}"];
                labourVarRowCopy2.Copy(Type.Missing);
                labourVarRowPaste2.PasteSpecial(Excel.XlPasteType.xlPasteValues);
                Excel.Range labourVarRowCopy3 = summarySheet.Range["J4:O4"];
                Excel.Range labourVarRowPaste3 = labourVarTrendSheet1.Range[$"K{labourVarTrendLastRow + 3} : P{labourVarTrendLastRow + 3}"];
                labourVarRowCopy3.Copy(Type.Missing);
                labourVarRowPaste3.PasteSpecial(Excel.XlPasteType.xlPasteValues);

                labourVarTrendSheet2.Rows[2].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                Excel.Range labourVarCopyRange2 = summarySheet.Range["B4:F4"];
                Excel.Range labourVarPasteRange2 = labourVarTrendSheet2.Range["A2:E2"];
                Excel.Range labourVarCopyRange3 = labourVarTrendSheet2.Range["A3:E3"];


                labourVarCopyRange3.Copy(Type.Missing);
                labourVarPasteRange2.PasteSpecial(Excel.XlPasteType.xlPasteFormats);

                labourVarCopyRange2.Copy(Type.Missing);
                labourVarPasteRange2.PasteSpecial(Excel.XlPasteType.xlPasteValues);

                labourVarTrendSheet2.Cells[2, 1].Value = date;

                Excel.ChartObjects chartObjects = labourVarTrendSheet2.ChartObjects() as Excel.ChartObjects;
                Excel.ChartObject chartObject2 = chartObjects.Item(2);
                Excel.Chart chart2 = chartObject2.Chart;

                int lastRow = labourVarTrendSheet2.UsedRange.Rows.Count;
                Excel.Range rangeA = labourVarTrendSheet2.Range["A2:A" + lastRow];

                Excel.Range rangeE = labourVarTrendSheet2.Range["E2:E" + lastRow];

                Excel.Range categoryRange = rangeA;
                Excel.Range valueRange = rangeE;

                Excel.Range combinedRange = labourVarTrendSheet2.Application.Union(categoryRange, valueRange);

                chart2.SetSourceData(combinedRange);



                Excel.ChartObject chartObject1 = chartObjects.Item(1);
                Excel.Chart chart1 = chartObject1.Chart;

                Excel.Range rangeB = labourVarTrendSheet2.Range["B2:B" + lastRow];

                Excel.Range valueRange2 = rangeB;

                Excel.Range combinedRange2 = labourVarTrendSheet2.Application.Union(categoryRange, valueRange2);

                chart1.SetSourceData(combinedRange2);

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

    internal class CjLabourStandardData
    {
        public string CjStore { get; set; }
        public double Sales { get; set; }
        public string Actual { get; set; }
    }

    internal class FfcglData
    {
        public string Entity { get; set; }
        public string Store { get; set; }
        public string JobDesp { get; set; }
        public string Desc { get; set; }
        public string Amount { get; set; }
    }
}
