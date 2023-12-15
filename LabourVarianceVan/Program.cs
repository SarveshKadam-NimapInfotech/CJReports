using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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

                Worksheet labourVar1SalesSheet = labourVar1Workbook.Worksheets["Sales"];

                int salesColumn = labourVar1SalesSheet.UsedRange.Columns.Count - 1;

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

                labourVar1FfcglSheet.Range[$"A{labourVar1FfcglSheetLastRow2}:A{labourVar1FfcglSheetLastRow1}"].Value = date;

                Worksheet pivotSheet = labourVar1Workbook.Worksheets["Data"];
                PivotTable pivotTable = pivotSheet.PivotTables(1);
                pivotTable.RefreshTable();

                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                string month = parsedDate.ToString("MMM", CultureInfo.InvariantCulture);
                string day = parsedDate.Day.ToString();

                var pivotDate = $"{day}-{month}";

                Excel.Range pivotRange = pivotSheet.PivotTables(1).TableRange1;

                foreach (Excel.Range cell in pivotRange.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString() == pivotDate)
                    {
                        Console.WriteLine($"Found date '{pivotDate}' at cell: {cell.Address}");
                        break; // Assuming you want to find only the first occurrence
                    }
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
}
