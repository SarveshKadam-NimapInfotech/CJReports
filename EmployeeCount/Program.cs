
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
    internal class employeeCountRowData
    {
        public string Entity { get; set; }

        public string Store { get; set; }

        public string Name { get; set; }
    }
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
                targetSouthSheet.Cells[targetSouthSheet.UsedRange.Rows.Count, 2].Value = "";

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
                targetNorthSheet.Cells[6, 2].Value = "";

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
 
                List<employeeCountRowData> ffcchksList = new List<employeeCountRowData>();
                int ffcchjsLastRow = worksheet.Cells[worksheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row;

                for (int i = 1; i <= ffcchjsLastRow; i++)
                {
                    string entity = Convert.ToString(worksheet.Cells[i, 1].Value);
                    string store = Convert.ToString(worksheet.Cells[i, 2].Value);
                    string name = Convert.ToString(worksheet.Cells[i, 6].Value);

                    if (!string.IsNullOrWhiteSpace(entity) && (entity.Contains("DFG") || entity.Contains("FSH") || entity.Contains("NWSM") || entity.Contains("RCIH") || entity.Contains("SUN")))
                    {
                        if (!string.IsNullOrWhiteSpace(store) && store != "0")
                        {
                            employeeCountRowData rowData = new employeeCountRowData
                            {
                                Store = store,
                                Name = name
                            };

                            ffcchksList.Add(rowData);
                        }
                    }
                }

                int sheet1Row =  2;
                foreach (var data in ffcchksList)
                {
                    targetSheet1.Cells[sheet1Row, 1].Value = data.Store;
                    targetSheet1.Cells[sheet1Row, 2].Value = data.Name;
                    sheet1Row++;
                }

                Dictionary<string, int> storeCounts = new Dictionary<string, int>();

                foreach (var data in ffcchksList)
                {
                    if (storeCounts.ContainsKey(data.Store))
                    {
                        storeCounts[data.Store]++;
                    }
                    else
                    {
                        storeCounts[data.Store] = 1;
                    }
                }

                Excel.Range clearRange = targetSheet1.Range[targetSheet1.Cells[5, 7], targetSheet1.Cells[targetSheet1.UsedRange.Rows.Count,7]];
                clearRange.Clear();

                int targetRow = 5;
                foreach (var pair in storeCounts)
                {
                    targetSheet1.Cells[targetRow, 6].Value = pair.Key;
                    targetSheet1.Cells[targetRow, 7].Value = pair.Value;
                    targetRow++;
                }
                targetSheet1.Cells[targetRow, 6].Value = "Grand Total";

                int total = storeCounts.Values.Sum();
                targetSheet1.Cells[targetRow, 7].Value = total;

                PivotTable pivotTable = targetSheet1.PivotTables(1);
                pivotTable.RefreshTable();

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
