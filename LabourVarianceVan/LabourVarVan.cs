using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using Excel = Microsoft.Office.Interop.Excel;
using Calendar = System.Globalization.Calendar;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml.Drawing;

namespace LabourVarianceVan
{
    internal class LabourVarVan
    {
        public static void CopySheetData(ExcelWorksheet sourceSheet, ExcelWorksheet destinationSheet)
        {
            // Get the dimensions of the source sheet
            var startCell = sourceSheet.Cells.Start.Row;
            var lastRow = sourceSheet.Cells[sourceSheet.Dimension.Start.Row, 1, sourceSheet.Dimension.End.Row, 1]
                                    .Reverse()
                                    .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
                                    ?.Start.Row ?? 0;
            var startColumn = sourceSheet.Cells.Start.Column;
            var endColumn = sourceSheet.Dimension.End.Column;


            // Iterate through each cell in the source sheet
            for (int row = startCell; row <= lastRow; row++)
            {
                for (int col = startColumn; col <= endColumn; col++)
                {
                    // Get the cell value from the source sheet
                    var cellValue = sourceSheet.Cells[row, col].Text;

                    // Set the cell value in the destination sheet
                    destinationSheet.Cells[row, col].Value = cellValue;
                }
            }
        }
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
        internal void LabourVarVanEpPlus()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string storeListFilePath = @"C:\Users\Public\Documents\StoreList.xlsx";

            string labourVar1FilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\Labor Var 2023-12-04.xlsx";

            string labourVar2FilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\Labor Var 2023.12.04.xlsx";

            string ffcglFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\FFCGL check date 12.25.23.xlsx";

            string cjLabourFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\CJ Labor Standard 2023-12-04.xlsx";

            string labourVarTrendFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\Labor Var Trend since 8.29.22.xlsx";

            string weeklySalesFilePath = @"C:\Users\Nimap\Downloads\Labor var Van - Copy\Week 51 Sales 2023-12-18.xlsm";

            var date = "12/18/2023";

            DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            string month = parsedDate.ToString("MMM", CultureInfo.InvariantCulture);
            string day = parsedDate.Day.ToString();

            var pivotDate = $"{day}-{month}";

            string cjMonth = parsedDate.ToString("MM", CultureInfo.InvariantCulture);
            string cjDay = parsedDate.ToString("dd", CultureInfo.InvariantCulture);
            string cjYear = parsedDate.ToString("yyyy", CultureInfo.InvariantCulture);

            var cjDate = $"{cjMonth}{cjDay}";

            DateTime twoWeeksPreviousDate = parsedDate.AddDays(-14);

            string previousCjMonth = twoWeeksPreviousDate.ToString("MM", CultureInfo.InvariantCulture);
            string previousCjday = twoWeeksPreviousDate.ToString("dd", CultureInfo.InvariantCulture);
            string previousCjYear = twoWeeksPreviousDate.ToString("yyyy", CultureInfo.InvariantCulture);

            var cjTwoWeeksPreviousDate = $"{previousCjMonth}{previousCjday}";

            List<double> firstValue = new List<double>();
            List<double> secondValue = new List<double>();
            List<CjLabourStandardData> cjLabourStandardDatas = new List<CjLabourStandardData>();



            FileInfo weeklySalesFile = new FileInfo(weeklySalesFilePath);
            FileInfo cjLabourFile = new FileInfo(cjLabourFilePath);
            FileInfo labourVar1File = new FileInfo(labourVar1FilePath);
            FileInfo ffcglFile = new FileInfo(ffcglFilePath);
            string modifiedLabourVar1 = $@"C:\Users\Public\Documents\test_{Guid.NewGuid()}.xlsx"; ;

            using (ExcelPackage weeklySalesPackage = new ExcelPackage(weeklySalesFile))
            using (ExcelPackage cjLabourPackage = new ExcelPackage(cjLabourFile))
            using (ExcelPackage labourVar1Package = new ExcelPackage(labourVar1File))
            using (ExcelPackage ffcglPackage = new ExcelPackage(ffcglFile))
            {

                // Get current week number and previous week number
                DateTime dateValue;
                DateTime.TryParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateValue);

                Calendar cal = CultureInfo.CurrentCulture.Calendar;
                var currentWeekNbr = cal.GetWeekOfYear(dateValue, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                var previousWeekNbr = currentWeekNbr - 1;

                int insertIndex = 1; // Start from the first sheet by default
                foreach (ExcelWorksheet sheet in labourVar1Package.Workbook.Worksheets)
                {
                    if (sheet.Name.StartsWith("Week"))
                    {
                        insertIndex = sheet.Index + 1;
                    }
                }

                // Add the new sheets after the last matching sheet
                ExcelWorksheet newPreviousWeekSheet = labourVar1Package.Workbook.Worksheets.Add($"Week {previousWeekNbr}");
                ExcelWorksheet newCurrentWeekSheet = labourVar1Package.Workbook.Worksheets.Add($"Week {currentWeekNbr}");

                // Move the newly added sheets to the desired position
                labourVar1Package.Workbook.Worksheets.MoveAfter(newPreviousWeekSheet.Name, labourVar1Package.Workbook.Worksheets[insertIndex - 1].Name);
                labourVar1Package.Workbook.Worksheets.MoveAfter(newCurrentWeekSheet.Name, newPreviousWeekSheet.Name);


                foreach (ExcelWorksheet worksheet in weeklySalesPackage.Workbook.Worksheets)
                {
                    string sheetName = worksheet.Name;
                    if (sheetName.StartsWith("Week"))
                    {
                        int weekNbr;
                        if (int.TryParse(sheetName.Split(' ')[1], out weekNbr))
                        {
                            if (weekNbr == currentWeekNbr || weekNbr == previousWeekNbr)
                            {
                                ExcelWorksheet lastWeekSheet = null;
                                ExcelWorksheet previousWeekSheet = null;

                                foreach (ExcelWorksheet sheet in labourVar1Package.Workbook.Worksheets)
                                {

                                    if (sheet.Name.StartsWith($"Week {currentWeekNbr}"))
                                    {
                                        lastWeekSheet = sheet;
                                    }
                                    if (sheet.Name.StartsWith($"Week {previousWeekNbr}"))
                                    {
                                        previousWeekSheet = sheet;
                                    }

                                }
                                if (lastWeekSheet != null && weekNbr == currentWeekNbr)
                                {
                                    var worksheetLastRow = worksheet.Cells[worksheet.Dimension.Start.Row, 1, worksheet.Dimension.End.Row, 1]
                                    .Reverse()
                                    .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
                                    ?.Start.Row ?? 0;

                                    ExcelRangeBase sourceRange = worksheet.Cells[1, 1, worksheetLastRow, worksheet.Dimension.End.Column];
                                    ExcelRangeBase destinationRange = lastWeekSheet.Cells[1, 1, worksheetLastRow, worksheet.Dimension.End.Column];


                                    sourceRange.Copy(destinationRange);
                                    destinationRange.Copy(lastWeekSheet.Cells[1, 1]);



                                }
                                if (previousWeekSheet != null && weekNbr == previousWeekNbr)
                                {


                                    var worksheetLastRow = worksheet.Cells[worksheet.Dimension.Start.Row, 1, worksheet.Dimension.End.Row, 1]
                                    .Reverse()
                                    .FirstOrDefault(cell => !string.IsNullOrWhiteSpace(cell.Text))
                                    ?.Start.Row ?? 0;

                                    ExcelRangeBase sourceRange = worksheet.Cells[1, 1, worksheetLastRow, worksheet.Dimension.End.Column];
                                    ExcelRangeBase destinationRange = previousWeekSheet.Cells[1, 1, worksheetLastRow, worksheet.Dimension.End.Column];


                                    sourceRange.Copy(destinationRange);
                                    destinationRange.Copy(lastWeekSheet.Cells[1, 1]);



                                }
                                var imagePath = @"C:\Users\Public\Documents\cjstar.png";
                                using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
                                {
                                    ExcelCalculationOption excelCalculation = new ExcelCalculationOption()
                                    {
                                        AllowCircularReferences = true
                                    };
                                    using (ExcelPicture picture = lastWeekSheet.Drawings.AddPicture("cjStar" + Guid.NewGuid().ToString(), stream))
                                    {
                                        picture.SetPosition(0, 0, 0, 0);
                                        picture.SetSize(130, 100);
                                    }
                                    using (ExcelPicture picture = previousWeekSheet.Drawings.AddPicture("cjStar" + Guid.NewGuid().ToString(), stream))
                                    {
                                        picture.SetPosition(0, 0, 0, 0);
                                        picture.SetSize(130, 100);
                                    }
                                    lastWeekSheet.Columns.AutoFit();
                                    lastWeekSheet.View.FreezePanes(7, 1);
                                    previousWeekSheet.Columns.AutoFit();
                                    previousWeekSheet.View.FreezePanes(7, 1);
                                }
                            }
                        }
                    }
                }

                labourVar1Package.Save();


                ExcelWorksheet labourVar1SalesSheet = labourVar1Package.Workbook.Worksheets["Sales"];

                int salesColumn = labourVar1SalesSheet.Dimension.Columns - 1;

                string salesColumnLetter = GetExcelColumnName(salesColumn);

                ExcelRangeBase sourceColumn = labourVar1SalesSheet.Cells[1, salesColumn, labourVar1SalesSheet.Dimension.End.Row, salesColumn];
                ExcelRangeBase destinationColumn = labourVar1SalesSheet.Cells[1, salesColumn + 1, labourVar1SalesSheet.Dimension.End.Row, salesColumn + 1];

                sourceColumn.Copy(destinationColumn);
                destinationColumn.AutoFitColumns();
                sourceColumn.ClearFormulaValues();
                sourceColumn.AutoFitColumns();

                //labourVar1SalesSheet.InsertColumn(salesColumn, 1);

                labourVar1SalesSheet.Cells[3, salesColumn].Value = date;
                labourVar1SalesSheet.Cells[4, salesColumn].Formula = $"=SUM({salesColumnLetter}5:{salesColumnLetter}{labourVar1SalesSheet.Dimension.Rows + 1})";

                string formula = $"=IFERROR((VLOOKUP($B5,'Week {previousWeekNbr}'!$A:$C,3,FALSE)+VLOOKUP($B5,'Week {currentWeekNbr}'!$A:$C,3,FALSE)),0)";
                labourVar1SalesSheet.Cells[$"{salesColumnLetter}5:{salesColumnLetter}{labourVar1SalesSheet.Dimension.Rows + 1}"].Formula = formula;

                labourVar1Package.Save();


                ExcelWorksheet ffcglsheet = ffcglPackage.Workbook.Worksheets["FFCGL"];
                List<FfcglData> ffcglList = new List<FfcglData>();
                int ffcglLastRow = ffcglsheet.Dimension.End.Row;

                for (int i = 1; i <= ffcglLastRow; i++)
                {
                    string entity = Convert.ToString(ffcglsheet.Cells[i, 1].Value);
                    string store = Convert.ToString(ffcglsheet.Cells[i, 2].Value);
                    string jobDesp = Convert.ToString(ffcglsheet.Cells[i, 3].Value);
                    string desc = Convert.ToString(ffcglsheet.Cells[i, 4].Value);
                    string amount = Convert.ToString(ffcglsheet.Cells[i, 5].Value);

                    if (!string.IsNullOrWhiteSpace(entity) && (entity.Contains("DFG") || entity.Contains("FSH") || entity.Contains("SUN")))
                    {
                        if (!string.IsNullOrWhiteSpace(desc) && desc.StartsWith("E") && !desc.Contains("Bonus"))
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

                // Access the "FFCGL" worksheet
                ExcelWorksheet labourVar1FfcglSheet = labourVar1Package.Workbook.Worksheets["FFCGL"];

                // Find the last row in column 1 of the "FFCGL" worksheet
                int labourVar1FfcglSheetLastRow1 = labourVar1FfcglSheet.Cells[labourVar1FfcglSheet.Dimension.End.Row, 1].End.Row + 1;

                int labourVar1FfcglSheetLastRow2 = labourVar1FfcglSheetLastRow1;


                foreach (var data in ffcglList)
                {
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 2].Value = data.Entity;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 3].Value = data.Store;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 4].Value = data.JobDesp;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 5].Value = data.Desc;
                    labourVar1FfcglSheet.Cells[labourVar1FfcglSheetLastRow1, 6].Value = data.Amount;
                    labourVar1FfcglSheetLastRow1++;
                }

                labourVar1FfcglSheet.Cells[$"A{labourVar1FfcglSheetLastRow2}:A{labourVar1FfcglSheetLastRow1 - 1}"].Value = date;

                labourVar1Package.Save();

                ExcelWorksheet pivotSheet = labourVar1Package.Workbook.Worksheets["Data"];
                var pivotTable = pivotSheet.PivotTables[0];

                pivotTable.CacheDefinition.SourceRange = labourVar1FfcglSheet.Cells[1, 1, labourVar1FfcglSheet.Dimension.End.Row, labourVar1FfcglSheet.Dimension.End.Column];



                var rowField = pivotTable.RowFields;
                foreach (var item in rowField)
                {
                    item.Items.Refresh();
                }
                var columnField = pivotTable.ColumnFields;
                foreach (var item in columnField)
                {
                    item.Items.Refresh();
                }
                labourVar1Package.Workbook.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });


                labourVar1Package.Save();

                string pivotColumnLetter = string.Empty;


                for (int i = 1; i < pivotSheet.Dimension.End.Column; i++)
                {
                    if (pivotSheet.Cells[1, i].Value == null)
                    {
                        pivotColumnLetter = GetExcelColumnName(i);
                        break;
                    }
                }



                //ExcelWorksheet pivotSheet = labourVar1Package.Workbook.Worksheets["Data"];

                //string pivotColumnLetter = string.Empty;
                //for (int i = 1; i < pivotSheet.Dimension.End.Column; i++)
                //{
                //    if (pivotSheet.Cells[1,i].Value == null)
                //    {
                //        pivotColumnLetter = HelperMethod.GetExcelColumnName(i);
                //        break;
                //    }
                //}


                //var pivotTable = pivotSheet.PivotTables.FirstOrDefault(); // Assuming there's only one PivotTable

                //string pivotColumnLetter = string.Empty;

                //if (pivotTable != null)
                //{
                //    // Get the data source range of the pivot table
                //    var dataSourceRange = pivotTable.CacheDefinition.SourceRange;

                //    // Calculate the range of the pivot table based on its data source
                //    var pivotRange = pivotSheet.Cells[Convert.ToInt32(dataSourceRange.Start.Address), Convert.ToInt32(dataSourceRange.End.Address)];

                //    // Use pivotRange as needed

                //    foreach (Excel.Range cell in pivotRange)
                //    {
                //        if (cell.Value != null && cell.Value.ToString() == pivotDate)
                //        {
                //            string cellAddress = cell.Address.ToString();
                //            pivotColumnLetter = new string(cellAddress.Where(char.IsLetter).ToArray());
                //            break;
                //        }
                //    }
                //}

                ExcelWorksheet labourVar1LaborsSheet = labourVar1Package.Workbook.Worksheets["Labors"];

                int laborsColumn = labourVar1LaborsSheet.Dimension.Columns - 1; // Getting the last column index

                string laborsColumnLetter = GetExcelColumnName(laborsColumn);

                // Insert a column in EPPlus
                labourVar1LaborsSheet.InsertColumn(laborsColumn + 1, 1);

                labourVar1LaborsSheet.Cells[3, laborsColumn + 1].Value = date;
                labourVar1LaborsSheet.Cells[4, laborsColumn + 1].Formula = $"=SUM({laborsColumnLetter}5:{laborsColumnLetter}60)";

                // Set formula for the range in EPPlus
                labourVar1LaborsSheet.Cells[$"{laborsColumnLetter}5:{laborsColumnLetter}60"].Formula = $"=INDEX(Data!{pivotColumnLetter}:{pivotColumnLetter},MATCH($B5,Data!$A:$A,0))";

                labourVar1Package.Save();

                ExcelWorksheet previousSheet = null;
                foreach (ExcelWorksheet sheet in cjLabourPackage.Workbook.Worksheets)
                {
                    if (sheet.Name == $"Labor Standard Var {cjTwoWeeksPreviousDate}")
                    {
                        previousSheet = sheet;
                        break;
                    }
                }


                ExcelWorksheet newSheet = cjLabourPackage.Workbook.Worksheets.Add($"Labor Standard Var {cjDate}", previousSheet);

                ExcelRange cellC4 = newSheet.Cells["C4"];
                cellC4.Value = date;

                ExcelWorksheet cjLabourSalesSheet = cjLabourPackage.Workbook.Worksheets["Sales"];
                var cjPreviousLabourDate = $"PPE {previousCjMonth}/{previousCjday}";

                cjLabourPackage.Save();

                ExcelRange row2 = cjLabourSalesSheet.Cells["2:2"];
                int cjLabourColumn = -1;
                foreach (ExcelRangeBase cell in row2)
                {
                    if (cell.Text == cjPreviousLabourDate)
                    {
                        cjLabourColumn = cell.Start.Column;
                        break;
                    }
                }

                string cjLaborsColumnLetter = GetExcelColumnName(cjLabourColumn + 1);

                cjLabourSalesSheet.InsertColumn(cjLabourColumn + 1, 1);

                cjLabourSalesSheet.Cells[2, cjLabourColumn + 1].Value = $"PPE {cjMonth}/{cjDay}";
                cjLabourSalesSheet.Cells[3, cjLabourColumn + 1].Value = "$";

                ExcelRange sourceLabourRange = labourVar1LaborsSheet.Cells[$"{laborsColumnLetter}5:{laborsColumnLetter}{labourVar1LaborsSheet.Dimension.Rows}"];
                ExcelRange targetLabourRange = cjLabourSalesSheet.Cells[$"{cjLaborsColumnLetter}5:{cjLaborsColumnLetter}{cjLabourSalesSheet.Dimension.Rows}"];

                // Loop through source range and set values in the target range
                for (int i = 1; i <= sourceLabourRange.Rows; i++)
                {
                    for (int col = 1; col <= sourceLabourRange.Columns; col++)
                    {
                        var sourceCell = sourceLabourRange[i, col];
                        var targetCell = targetLabourRange[i, col];

                        targetCell.Value = sourceCell.Value;
                    }
                }

                cjLabourSalesSheet.Cells[4, cjLabourColumn + 1].Formula = $"=SUM({cjLaborsColumnLetter}5:{cjLaborsColumnLetter}60)";

                var cjPreviousSalesDate = $"Ending {previousCjMonth}/{previousCjday}";

                ExcelRange row3 = cjLabourSalesSheet.Cells["3:3"];
                int cjSalesColumn = -1;
                foreach (var cell in row3)
                {
                    if (cell.Text == $"Ending {cjMonth}/{cjDay}")
                    {
                        cjSalesColumn = cell.Start.Column;
                        break;
                    }
                }

                string cjSalesColumnLetter = GetExcelColumnName(cjSalesColumn);

                cjLabourSalesSheet.InsertColumn(cjSalesColumn + 1, 1);

                cjLabourSalesSheet.Cells[3, cjSalesColumn + 1].Value = $"Ending {cjMonth}/{cjDay}";

                var sourceSalesRange = labourVar1SalesSheet.Cells[$"{salesColumnLetter}5:{salesColumnLetter}{labourVar1SalesSheet.Dimension.Rows}"];
                var targetSalesRange = cjLabourSalesSheet.Cells[$"{cjSalesColumnLetter}5:{cjSalesColumnLetter}{cjLabourSalesSheet.Dimension.Rows}"];

                targetSalesRange.LoadFromCollection(sourceSalesRange, false); // Copy values from source to target

                cjLabourSalesSheet.Cells[4, cjSalesColumn + 1].Formula = $"=SUM({cjSalesColumnLetter}5:{cjSalesColumnLetter}60)";

                newSheet.Cells["C9:C62"].Formula = $"=INDEX(Sales!{cjSalesColumnLetter}:{cjSalesColumnLetter},MATCH('Labor Standard Var 1204'!B9,Sales!B:B,0))";
                newSheet.Cells["F9:F62"].Formula = $"=INDEX(Sales!{cjLaborsColumnLetter}:{cjLaborsColumnLetter},MATCH('Labor Standard Var 1204'!B9,Sales!B:B,0))";

                ExcelWorksheet cjLabourStandardSheet = cjLabourPackage.Workbook.Worksheets[$"Labor Standard Var {cjDate}"];

                int cjLabourStandardLastRow = 0;

                // Finding the last row based on non-null values in column B
                for (int rows = 9; rows <= cjLabourStandardSheet.Dimension.End.Row; rows++)
                {
                    if (cjLabourStandardSheet.Cells[rows, 2].Value == null)
                    {
                        cjLabourStandardLastRow = rows;
                        break;
                    }
                }


                // Iterating through the rows to collect data
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

                // Accessing the "Changing list" worksheet
                ExcelWorksheet changingListSheet = cjLabourPackage.Workbook.Worksheets["Changing list"];

                List<string> regList = new List<string>();

                // Gathering data from column 13 of "Changing list" worksheet
                for (int i = 9; i <= cjLabourStandardLastRow; i++)
                {
                    regList.Add(Convert.ToString(changingListSheet.Cells[i, 13].Value));
                }

                var changeListRow = 9;
                // Writing data from regList to "Labor Standard Var {cjDate}" worksheet
                foreach (var data in regList)
                {
                    cjLabourStandardSheet.Cells[changeListRow, 13].Value = data;
                    changeListRow++;
                }

                int lastRowUsed = 0;
                for (int rows = 9; rows < cjLabourStandardSheet.Dimension.End.Row; rows++)
                {
                    if (cjLabourStandardSheet.Cells[rows, 5].Value == null)
                    {
                        lastRowUsed = rows;
                        break;
                    }

                }

                ExcelRangeBase range1 = cjLabourStandardSheet.Cells[$"E9:E{lastRowUsed - 1}"];
                ExcelRangeBase range2 = cjLabourStandardSheet.Cells[$"N9:N{lastRowUsed - 1}"];


                range1.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });
                range2.Calculate(new ExcelCalculationOption() { AllowCircularReferences = true });


                for (int rows = 9; rows <= lastRowUsed; rows++)
                {
                    if (cjLabourStandardSheet.Cells[rows, 5].Value != null && cjLabourStandardSheet.Cells[rows, 5].Value is double)
                    {
                        firstValue.Add(Convert.ToDouble(cjLabourStandardSheet.Cells[rows, 5].Value));
                    }
                    if (cjLabourStandardSheet.Cells[rows, 14].Value != null && cjLabourStandardSheet.Cells[rows, 14].Value is double)
                    {
                        secondValue.Add(Convert.ToDouble(cjLabourStandardSheet.Cells[rows, 14].Value));
                    }
                }

                labourVar1Package.Save();
                cjLabourPackage.Save();
                weeklySalesPackage.Save();

            }

        }
    }


        
    
}