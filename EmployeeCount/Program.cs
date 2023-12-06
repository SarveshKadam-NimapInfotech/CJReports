
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmployeeCount
{
    internal class Program
    {
        static void Main(string[] args)
        {
             EmployeeCount();

        }

         static void EmployeeCount()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string filePath = @"C:\Users\Nimap\Downloads\Employee count\FFCCHKS - Copy.xlsx";
            Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);

            string targetfilePath = @"C:\Users\Nimap\Downloads\Employee count\Employee counts - Copy.xlsx";
            Excel.Workbook targetWorkbook = excelApp.Workbooks.Open(targetfilePath);

            try
            {
                
                Worksheet targetSouthSheet = targetWorkbook.Worksheets["South"];

                Excel.Range southColumnB = targetSouthSheet.Columns["B:B"];
                southColumnB.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                Excel.Range southColumnBCleared = targetSouthSheet.Columns["B:B"];
                southColumnBCleared.Clear();

                targetSouthSheet.Range[$"B2:B{targetSouthSheet.UsedRange.Rows.Count}"].Formula = "=VLOOKUP(A2,Sheet1!F:G,2,0)";


                Excel.Range dateSouthCell = targetSouthSheet.Cells[1, 3];

                if (DateTime.TryParse(dateSouthCell.Text, out DateTime southCurrentDate))
                {
                    DateTime newDate = southCurrentDate.AddDays(14);

                    Excel.Range newDateCell = targetSouthSheet.Cells[1, 2]; 
                    newDateCell.NumberFormat = "MM/dd/yyyy";
                    newDateCell.Value = newDate;
                    newDateCell.Font.Bold = true;
                    newDateCell.EntireColumn.AutoFit();
                }

                Worksheet targetNorthSheet = targetWorkbook.Worksheets["North"];

                Excel.Range northColumnB = targetNorthSheet.Columns["B:B"];
                northColumnB.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                Excel.Range northColumnBCleared = targetNorthSheet.Columns["B:B"];
                northColumnBCleared.Clear();

                targetNorthSheet.Range[$"B2:B{targetNorthSheet.UsedRange.Rows.Count}"].Formula = "=VLOOKUP(A2,Sheet1!F:G,2,0)";

                Excel.Range dateNorthCell = targetNorthSheet.Cells[1, 3];

                if(DateTime.TryParse(dateNorthCell.Text, out DateTime northCurrentDate))
                {
                    DateTime newDate = northCurrentDate.AddDays(14);

                    Excel.Range newDateCell = (Excel.Range)targetNorthSheet.Cells[1, 2];
                    newDateCell.NumberFormat = "MM/dd/yyyy";
                    newDateCell.Value = newDate;
                    newDateCell.Font.Bold = true;
                    newDateCell.EntireColumn.AutoFit();
                }
                

                Worksheet worksheet = workbook.Worksheets["FFCCHKS"];
                Worksheet targetSheet1 = targetWorkbook.Worksheets["Sheet1"];

                Excel.Range worksheetColumnB = worksheet.Columns["B"];

                Dictionary<object, int> valueCounts = new Dictionary<object, int>();
                List<object> uniqueValues = new List<object>();

                int rowCount = worksheetColumnB.Rows.Count;
                int currentRow = 1;
                int nullCount = 0;
                const int maxNullCount = 10;

                while (currentRow <= rowCount)
                {
                    var cell = worksheetColumnB.Cells[currentRow, 1];
                    var value = cell.Value2;

                    if (value == null || value.ToString() == "0" || value.ToString() == "")
                    {
                        nullCount++;

                        if (nullCount > maxNullCount)
                        {
                            break;
                        }
                    }
                    else
                    {
                        nullCount = 0; 

                        if (!uniqueValues.Contains(value))
                        {
                            uniqueValues.Add(value);
                        }

                        if (valueCounts.ContainsKey(value))
                        {
                            valueCounts[value]++;
                        }
                        else
                        {
                            valueCounts[value] = 1;
                        }
                    }

                    currentRow++;
                }

                //foreach (var value in uniqueValues)
                //{
                //    Console.WriteLine($"Value: {value}, Count: {valueCounts[value]}");
                //}

                int targetRow = 5; 
                foreach (var value in uniqueValues)
                {
                    targetSheet1.Cells[targetRow, 6].Value = value; 
                    targetSheet1.Cells[targetRow, 7].Value = valueCounts[value]; 
                    targetRow++;
                }


                var filterEntity = new object[]
                {
                    "DFG",
                    "FSH",
                    "NWSM",
                    "RCIH",
                    "SUN"
                };

                var filterStore = new object[]
                {
                    //uniqueValues
                   
                };

                Range worksheetRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, worksheet.UsedRange.Column]];

                worksheetRange.AutoFilter(1, filterEntity, XlAutoFilterOperator.xlFilterValues, Type.Missing, true);
                var filterByEntity = worksheetRange.SpecialCells(XlCellType.xlCellTypeVisible);

                filterByEntity.AutoFilter(2, filterByEntity, XlAutoFilterOperator.xlFilterNoFill, Type.Missing, true);
                
                Excel.Range columnA = worksheet.Columns["B:B"];
                columnA.Copy(Type.Missing);



                Excel.Range columnB = worksheet.Columns["F:F"];
                columnB.Copy(Type.Missing);


                targetWorkbook.Save();
                targetWorkbook.Close();
                workbook.Close();
               


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
