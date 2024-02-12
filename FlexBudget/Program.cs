using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace FlexBudget
{
  
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.FlexBudget();
        }

        


        public void FlexBudget()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string flexBudgetFilePath = @"C:\Users\Nimap\Documents\FlexBudget\Flex Budget 2023-11 - CJNC.xlsx";

            string CjNorthFilePath = @"C:\Users\Nimap\Documents\FlexBudget\FF IS - CJ North 12.23.xlsx";

            Excel.Workbook flexBudgetWorkbook = excelApp.Workbooks.Open(flexBudgetFilePath);

            Excel.Workbook CjNorthWorkbook = excelApp.Workbooks.Open(CjNorthFilePath);

            try
            {
                var managementFessFormulaValue = 3.44;

                var date = "12/01/2023";

                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                int month = parsedDate.Month;

                int year = parsedDate.Year;

                int previousMonth = 0;
                if (month == 1)
                {
                    previousMonth = 12;
                }
                else
                {
                    previousMonth = month - 1;

                }

                string monthName = parsedDate.ToString("MMM", CultureInfo.InvariantCulture);

                Worksheet flexBudgetSheet = flexBudgetWorkbook.Worksheets["Flex Budget"];

                Worksheet uploadSheet = flexBudgetWorkbook.Worksheets["Upload (2)"];

                Worksheet uploadCleanSheet = flexBudgetWorkbook.Worksheets["Upload Clean"];




                switch (month)
                {
                    case 1:
                        if (previousMonth == 12)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"AA8:AB{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"E8:F{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"AA8:AB{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            //flexBudgetSheet.Cells[5, 5].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 217, 102));
                            //flexBudgetSheet.Cells[5, 6].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 217, 102));

                            //flexBudgetSheet.Range["E5"].Value = "Flex";
                            //flexBudgetSheet.Range["E6"].Value = "Flex";

                            //for(int col = 7; col <= 28; col++)
                            //{
                            //    flexBudgetSheet.Cells[5, col].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));
                            //    flexBudgetSheet.Cells[5, col].Value = "Fixed";
                            //}

                            //Range copyFlexRange = flexBudgetSheet.Range["E5:F5"];
                            //Range pasteFlexRange = flexBudgetSheet.Range["G5:H5"];

                            //copyFlexRange.Copy(Type.Missing);
                            //pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"E2:E{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"F2:F{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"E2:E{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["E"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["P"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "P"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "P"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "P"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "P"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"P4:P{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"P4:P{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"E4:E359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,4,FALSE)";


                        }
                        break;

                    case 2:
                        if (previousMonth == 1)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"E8:F{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"G8:H{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"E8:F{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["E5:F5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["G5:H5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"E2:E{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"F2:F{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"E2:E{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["F"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["E"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "E"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "E"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "E"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "E"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"E4:E{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"E4:E{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"F4:F359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,5,FALSE)";


                        }
                        break;

                    case 3:
                        if ( previousMonth == 2)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"G8:H{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"I8:J{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"G8:H{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["G5:H5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["I5:J5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"F2:F{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"G2:G{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"F2:F{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["G"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["F"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "F"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "F"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "F"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "F"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"F4:F{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"F4:F{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"G4:G359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,6,FALSE)";


                        }
                        break;

                    case 4:
                        if ( previousMonth == 3)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"I8:J{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"K8:L{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"I8:J{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["I5:J5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["K5:L5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"G2:G{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"H2:H{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"G2:G{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["H"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["G"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "G"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "G"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "G"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "G"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"G4:G{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"G4:G{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"H4:H359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,7,FALSE)";



                        }
                        break;

                    case 5:
                        if ( previousMonth == 4)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"K8:L{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"M8:N{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"K8:L{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["K5:L5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["M5:N5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"H2:H{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"I2:I{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"H2:H{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["I"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["H"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "H"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "H"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "H"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "H"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"H4:H{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"H4:H{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"I4:I359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,8,FALSE)";


                        }
                        break;

                    case 6:
                        if ( previousMonth == 5)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"M8:N{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"O8:P{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"M8:N{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["M5:N5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["O5:P5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"I2:I{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"J2:J{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"I2:I{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["J"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["I"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "I"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "I"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "I"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "I"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"I4:I{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"I4:I{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"J4:J359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,9,FALSE)";


                        }
                        break;

                    case 7:
                        if ( previousMonth == 6)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"O8:P{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"Q8:R{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"O8:P{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["O5:P5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["Q5:R5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"J2:J{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"K2:K{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"J2:J{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["K"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["J"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "J"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "J"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "J"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "J"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"J4:J{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"J4:J{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"K4:K359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,10,FALSE)";

                        }
                        break;

                    case 8:
                        if ( previousMonth == 7)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"Q8:R{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"S8:T{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"Q8:R{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["Q5:R5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["S5:T5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"K2:K{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"L2:L{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"K2:K{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["L"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["K"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "K"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "K"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "K"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "K"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"K4:K{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"K4:K{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"L4:L359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,11,FALSE)";

                        }
                        break;

                    case 9:
                        if ( previousMonth == 8)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"S8:T{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"U8:V{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"S8:T{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["S5:T5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["U5:V5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"L2:L{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"M2:M{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"L2:L{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["M"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["L"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "L"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "L"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "L"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "L"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"L4:L{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"L4:L{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"M4:M359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,12,FALSE)";


                        }
                        break;

                    case 10:
                        if ( previousMonth == 9)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"U8:V{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"W8:X{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"U8:V{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["U5:V5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["W5:X5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"M2:M{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"N2:N{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"M2:M{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["N"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["M"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "M"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "M"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "M"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "M"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"M4:M{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"M4:M{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"N4:N359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,13,FALSE)";


                        }
                        break;

                    case 11:
                        if ( previousMonth == 10)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"W8:X{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"Y8:Z{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"W8:X{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["W5:X5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["Y5:Z5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                           {
                                "Management Fee Expense",
                                "Delivery Charges"

                           };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"N2:N{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"O2:O{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"N2:N{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas);
                            //valuePasteRange2.PasteSpecial(XlPasteType.xlPasteValues);

                            Range filteredRangePaste = uploadSheet.Columns["O"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["N"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1;
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "N"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "N"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "N"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "N"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"N4:N{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"N4:N{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"O4:O359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,14,FALSE)";


                        }
                        break;

                    case 12:
                        if ( previousMonth == 11)
                        {
                            Range copyRange1 = flexBudgetSheet.Range[$"Y8:Z{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range pasteRange1 = flexBudgetSheet.Range[$"AA8:AB{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange = flexBudgetSheet.Range[$"Y8:Z{flexBudgetSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            valuePasteRange.PasteSpecial(XlPasteType.xlPasteValues);

                            Range copyFlexRange = flexBudgetSheet.Range["Y5:Z5"];
                            Range pasteFlexRange = flexBudgetSheet.Range["AA5:AB5"];

                            copyFlexRange.Copy(Type.Missing);
                            pasteFlexRange.PasteSpecial(XlPasteType.xlPasteAll);

                            // upload sheet task

                            var uploadSheetFilterList = new object[]
                            {
                                "Management Fee Expense",
                                "Delivery Charges"

                            };
                            Range sourceRange = uploadSheet.Range[uploadSheet.Cells[1, 1], uploadSheet.Cells[1, uploadSheet.UsedRange.Column]];
                            sourceRange.AutoFilter(3, uploadSheetFilterList, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);

                            //Range copyRange2 = uploadSheet.Range[$"O2:O{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range pasteRange2 = uploadSheet.Range[$"P2:P{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            //Range valuePasteRange2 = uploadSheet.Range[$"O2:O{uploadSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            //copyRange2.Copy(Type.Missing);
                            //pasteRange2.PasteSpecial(XlPasteType.xlPasteFormulas, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);
                            //copyRange2.PasteSpecial(XlPasteType.xlPasteValues, XlPasteSpecialOperation.xlPasteSpecialOperationNone, Type.Missing, Type.Missing);

                            Range filteredRangePaste = uploadSheet.Columns["P"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);
                            Range filteredRangeCopy = uploadSheet.Columns["O"].SpecialCells(Excel.XlCellType.xlCellTypeVisible);

                            int i = 1; 
                            foreach (Range cell in filteredRangePaste)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "O"].Value;

                                if(filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Formula = uploadSheet.Cells[cell.Row, "O"].Formula;
                                    i++;
                                }

                            }
                            foreach (Range cell in filteredRangeCopy)
                            {
                                var filteredValue = uploadSheet.Cells[cell.Row, "O"].Value;

                                if (filteredValue == null)
                                {
                                    break;
                                }
                                if (!filteredValue.ToString().Contains("Temp"))
                                {
                                    cell.Value = uploadSheet.Cells[cell.Row, "O"].Value;
                                    i++;
                                }

                            }

                            //uploadClean sheet task

                            Range copyRange3 = uploadCleanSheet.Range[$"O4:O{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];
                            Range valuePasteRange3 = uploadCleanSheet.Range[$"O4:O{uploadCleanSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row}"];

                            copyRange3.Copy(Type.Missing);
                            valuePasteRange3.PasteSpecial(XlPasteType.xlPasteValues);

                            uploadCleanSheet.Range[$"P4:P359"].Formula = "=VLOOKUP(B4,'Upload (2)'!$B$22:$P$1600,15,FALSE)";

                        }
                        break;

                    default:
                        break;
                }


                Worksheet RawDataPlSheet = flexBudgetWorkbook.Worksheets["Raw Data PL"];

                Worksheet CjNorthSheet = CjNorthWorkbook.Worksheets[1];

                Range clearRange = RawDataPlSheet.Range["B1:S" + RawDataPlSheet.Rows.Count];
                clearRange.Clear();

                Range copyRange = CjNorthSheet.Range["A1:R" + CjNorthSheet.Rows.Count];
                copyRange.Copy(Type.Missing);

                Range pasteRange = RawDataPlSheet.Range["B1:S" + RawDataPlSheet.Rows.Count];
                pasteRange.PasteSpecial(XlPasteType.xlPasteAll);

                Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

                int rawDataLastRow = RawDataPlSheet.Cells[RawDataPlSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 1;

                for (int i = 1; i <= rawDataLastRow; i++)
                {
                    var cellB = RawDataPlSheet.Cells[i, 2];
                    var cellC = RawDataPlSheet.Cells[i, 3];
                    var cellD = RawDataPlSheet.Cells[i, 4];

                    // Check if any of the cells are null
                    if (cellB.Value != null)
                    {
                        string cellValueB = Convert.ToString(cellB.Value);
                        string cellValueC = Convert.ToString(cellC?.Value);
                        string cellValueD = Convert.ToString(cellD?.Value);

                        if (cellValueB.Contains("Net Sales") || cellValueB.Contains("Labor Matrix") || cellValueB.Contains("Ideal Food Cost"))
                        {
                            string key = cellValueB;

                            if (!dict.ContainsKey(key))
                            {
                                dict[key] = new List<string>();
                            }

                            dict[key].Add(cellValueC ?? "");  // Value1
                            dict[key].Add(cellValueD ?? "");  // Value2
                        }
                    }
                }


                foreach (var kvp in dict)
                {
                    Console.WriteLine($"Key: {kvp.Key}");

                    List<string> values = kvp.Value;

                    for (int j = 0; j < values.Count; j += 2)
                    {
                        string value1 = values[j];
                        string value2 = (j + 1 < values.Count) ? values[j + 1] : "N/A";

                        Console.WriteLine($"  Value1: {value1}, Value2: {value2}");
                    }

                    Console.WriteLine();
                }

                Worksheet salesLoadSheet = flexBudgetWorkbook.Worksheets["Current Mo Sales Load"];

                Worksheet actualIdealSheet = flexBudgetWorkbook.Worksheets["Actual Ideal"];

                //Printing the values in Sales Load Sheet

                int startRowNetSales = 6; 
                int startRowLabourMatrix = 2; 
                int startRowIdealFoodCost = 19; 

                foreach (var kvp in dict)
                {
                    string key = kvp.Key;
                    List<string> values = kvp.Value;

                    if (key.Contains("Net Sales"))
                    {
                        for (int j = 0; j < values.Count; j += 2)
                        {
                            string value1 = values[j];
                            string value2 = (j + 1 < values.Count) ? values[j + 1] : "N/A";

                            salesLoadSheet.Cells[startRowNetSales, 23].Value = value1;
                            salesLoadSheet.Cells[startRowNetSales, 24].Value = value2; 

                            startRowNetSales++;
                        }
                    }
                    else if (key.Contains("Labor Matrix"))
                    {
                        for (int j = 0; j < values.Count; j += 2)
                        {
                            string value1 = values[j];
                            string value2 = (j + 1 < values.Count) ? values[j + 1] : "N/A";

                            switch (month)
                            {
                                case 1 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 5].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 6].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 2 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 7].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 8].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 3 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 9].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 10].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 4 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 11].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 12].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 5 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 13].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 14].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 6 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 15].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 16].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 7 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 17].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 18].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 8 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 19].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 20].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 9 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 21].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 22].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 10 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 23].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 24].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 11 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 25].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 26].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 12 :
                                    actualIdealSheet.Cells[startRowLabourMatrix, 27].Value = value1;
                                    actualIdealSheet.Cells[startRowLabourMatrix, 28].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                default:
                                    break;

                            }

                            startRowLabourMatrix++;
                        }
                    }
                    else if (key.Contains("Ideal Food Cost"))
                    {
                        for (int j = 0; j < values.Count; j += 2)
                        {
                            string value1 = values[j];
                            string value2 = (j + 1 < values.Count) ? values[j + 1] : "N/A";

                            switch (month)
                            {
                                case 1:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 5].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 6].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 2:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 7].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 8].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 3:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 9].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 10].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 4:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 11].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 12].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 5:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 13].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 14].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 6:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 15].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 16].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 7:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 17].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 18].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 8:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 19].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 20].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 9:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 21].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 22].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 10:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 23].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 24].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 11:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 25].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 26].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                case 12:
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 27].Value = value1;
                                    actualIdealSheet.Cells[startRowIdealFoodCost, 28].Value = (value2 == "N/A") ? "N/A" : (double.Parse(value2) * 100).ToString();
                                    break;

                                default:
                                    break;

                            }

                            startRowIdealFoodCost++;
                        }
                    }
                }


                string managementFees = string.Empty;

                for (int i = 1; i <= rawDataLastRow; i++)
                {
                    var cellB = RawDataPlSheet.Cells[i, 2];
                    var cellC = RawDataPlSheet.Cells[i, 3];

                    // Check if any of the cells are null
                    if (cellB.Value != null)
                    {
                        string cellValueB = Convert.ToString(cellB.Value);
                        string cellValueC = Convert.ToString(cellC?.Value);

                        if(cellValueB.Contains("Management Fees"))
                        {
                            managementFees = cellValueC;
                            break;
                        }

                    }
                }


                salesLoadSheet.Range["H17"].Value = managementFees;

                salesLoadSheet.Range["H4"].Value = managementFessFormulaValue;

                salesLoadSheet.Range["H7:H13"].Formula = $"=D7*$H$4%";

                var managementFessValueH19 = Convert.ToDouble(salesLoadSheet.Range["H19"].Value);
                salesLoadSheet.Range["H7:H13"].Formula = $"=D7*$H$4% {(managementFessValueH19 >= 0 ? "-" : "+")}{Math.Abs(managementFessValueH19)}";

                //salesLoadSheet.Range["H7:H13"].Formula = $"=D8*{managementFessFormulaValue}%+{managementFessValueH19}";


                var managementFessValueH18 = Convert.ToDouble(salesLoadSheet.Range["H18"].Value);
                salesLoadSheet.Range["H7"].Formula = $"=D7*$H$4% {(managementFessValueH19 >= 0 ? "-" : "+")}{Math.Abs(managementFessValueH19)} - {managementFessValueH18}";

                salesLoadSheet.Range["D3"].Value = monthName;
                salesLoadSheet.Range["D4"].Value = month;

                DateTime salesW5 = Convert.ToDateTime($"{monthName}-{year}");

                Range cellW5 = salesLoadSheet.Range["W5"];
                cellW5.Value = salesW5;

                cellW5.NumberFormat = "MMM yyyy";



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
