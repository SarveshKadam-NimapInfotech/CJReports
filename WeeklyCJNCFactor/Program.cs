using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace WeeklyCJNCFactor
{
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

            var weeklyCjncFactorFolderPath = @"C:\Users\Public\Documents\Weekly CJNC Factor\";

            var weeklySalesFilePath = @"C:\Users\Nimap\Documents\WeeklyCjncFactor\weeklyCjncFactor-1716928341768.xlsm";

            Excel.Workbook weeklySalesWorkbook = excelApp.Workbooks.Open(weeklySalesFilePath);

            try
            {
                var date = "06/03/2024";
                DateTime dateValue;
                DateTime.TryParseExact(date, "MM/dd/yyyy", new CultureInfo("en-US"), DateTimeStyles.None, out dateValue);
                Calendar cal = new CultureInfo("en-US").Calendar;
                cal = CultureInfo.CurrentCulture.Calendar;

                var currentWeekNbr = cal.GetWeekOfYear(dateValue, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                currentWeekNbr -= 2;

                //int previousWeekNbr;
                //if (currentWeekNbr == 1)
                //{
                //    // If the current week number is 1, get the last week of the previous year
                //    previousWeekNbr = cal.GetWeekOfYear(dateValue.AddYears(-1), CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
                //}
                //else
                //{
                //    // Otherwise, get the previous week of the current year
                //    previousWeekNbr = currentWeekNbr - 1;
                //}

                Worksheet weekSheet = weeklySalesWorkbook.Worksheets[$"Week {currentWeekNbr}"];

                var dateCellValue = weekSheet.Cells[3,3].Value.ToString();

                Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

                int weekSheetLastRow = weekSheet.Cells[weekSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 1;

                string companyGrowth = string.Empty;

                for (int i = 9; i <= weekSheetLastRow; i++)
                {
                    string store = Convert.ToString(weekSheet.Cells[i, 1].Value);
                    string generalManager = Convert.ToString(weekSheet.Cells[i, 2].Value);
                    string storeGrowth = Convert.ToString(weekSheet.Cells[i, 6].Value);

                    string key = string.IsNullOrEmpty(store) ? generalManager : store;

                    if (key == "Greg Funkhouser")
                    {
                        companyGrowth = storeGrowth;
                    }

                    if (!string.IsNullOrEmpty(key)) // Ensure the key is not null or empty
                    {
                        if (!dict.ContainsKey(key))
                        {
                            dict[key] = new List<string>();
                        }
                        dict[key].Add(generalManager);
                        dict[key].Add(storeGrowth);
                    }

                    // Additional handling for "Isidro Camacho"
                    if (generalManager.Contains("Isidro Camacho"))
                    {
                        string isidroKey = "Isidro Camacho";

                        if (!dict.ContainsKey(isidroKey))
                        {
                            dict[isidroKey] = new List<string>();
                        }
                        dict[isidroKey].Add(generalManager);
                        dict[isidroKey].Add(storeGrowth);
                    }
                }

                var fileNames = Directory.GetFiles(weeklyCjncFactorFolderPath);
                List<string> weeklyCjncFiles = new List<string>();
                foreach (var file in fileNames)
                {
                    weeklyCjncFiles.Add(file);
                }

                foreach (var key in dict.Keys)
                {
                    var matchingFile = key == "Isidro Camacho"
                    ? weeklyCjncFiles.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).StartsWith("Isidro Camacho"))
                    : weeklyCjncFiles.FirstOrDefault(f => Path.GetFileNameWithoutExtension(f).Contains(key));

                    if (matchingFile != null)
                    {
                        // Open the workbook
                        Excel.Workbook workbook = excelApp.Workbooks.Open(matchingFile);

                        Excel.Worksheet sheet1 = workbook.Sheets["Sheet1"];
                        Excel.Worksheet sheet2 = workbook.Sheets["Sheet2"];

                        string cellDateString = sheet1.Cells[6, 1].Value.ToString();
                        DateTime cellDate;
                        DateTime.TryParse(cellDateString, out cellDate);
                        DateTime givenDate;
                        DateTime.TryParse(date, out givenDate);


                        if (key == "Greg Funkhouser" || key == "Isidro Camacho")
                        {

                        if (cellDate.Date != givenDate.Date)
                        {
                            // sheet 1 code
                            Excel.Range row6 = sheet1.Rows[6];
                            row6.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                            sheet1.Cells[6, 1].Value = date;

                            Excel.Range row7 = sheet1.Rows[7];
                            row7.Copy(Type.Missing);
                            Excel.Range newRow6 = sheet1.Rows[6];
                            newRow6.PasteSpecial(XlPasteType.xlPasteFormats);

                            sheet1.Range["C6"].Formula = companyGrowth;
                            sheet1.Range["B6"].Formula = "=C6-C3";

                            // sheet 2 code
                            Excel.Range row3 = sheet2.Rows[3];
                            row3.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                            Excel.Range row4 = sheet2.Rows[4];
                            row4.Copy(Type.Missing);
                            Excel.Range newRow3 = sheet2.Rows[3];
                            newRow3.PasteSpecial(XlPasteType.xlPasteFormats);

                            Excel.Range row11 = sheet2.Rows[11];
                            row11.Delete();

                            sheet2.Range["A3"].Value = date;
                            sheet2.Range["B3"].Formula = "=Sheet1!$C6/100";

                        }
                        else
                        {
                            sheet1.Range["C6"].Formula = companyGrowth;

                        }

                    }
                    else
                    {
                        if (cellDate.Date != givenDate.Date)
                        {
                            // sheet 1 code
                            Excel.Range row6 = sheet1.Rows[6];
                            row6.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                            sheet1.Cells[6, 1].Value = date;


                            if (dict[key].Count > 1)
                            {
                                sheet1.Cells[6, 4].Value = dict[key][1];
                            }

                            Excel.Range row7 = sheet1.Rows[7];
                            row7.Copy(Type.Missing);
                            Excel.Range newRow6 = sheet1.Rows[6];
                            newRow6.PasteSpecial(XlPasteType.xlPasteFormats);

                            sheet1.Range["E6"].Formula = companyGrowth;
                            sheet1.Range["B6"].Formula = "=C6-C3";
                            sheet1.Range["C6"].Formula = "=D6-E6";

                            // sheet 2 code
                            Excel.Range row3 = sheet2.Rows[3];
                            row3.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                            Excel.Range row4 = sheet2.Rows[4];
                            row4.Copy(Type.Missing);
                            Excel.Range newRow3 = sheet2.Rows[3];
                            newRow3.PasteSpecial(XlPasteType.xlPasteFormats);

                            Excel.Range row11 = sheet2.Rows[11];
                            row11.Delete();

                            sheet2.Range["A3"].Value = date;
                            sheet2.Range["B3"].Formula = "=Sheet1!$E6/100";
                            sheet2.Range["C3"].Formula = "=Sheet1!$D6/100";

                                workbook.Save();



                        }
                        else
                        {
                            if (dict[key].Count > 1)
                            {
                                sheet1.Cells[6, 4].Value = dict[key][1];

                            }

                            sheet1.Range["E6"].Formula = companyGrowth;

                        }

                           
                    }


                        


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
