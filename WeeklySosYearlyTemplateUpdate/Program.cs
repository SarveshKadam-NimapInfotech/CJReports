using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;

using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace WeeklySosYearlyTemplateUpdate
{
    internal class Program
    {
        static void Main(string[] args)
        {

            //string date = "12/30/2024";

            //DateTime startDate;
            //DateTime.TryParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate);
            //Calendar cal = CultureInfo.CurrentCulture.Calendar;
            //int weekNumber = cal.GetWeekOfYear(startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            //// Determine the year of the given week number
            //int year = GetYearOfWeekNumber(startDate, weekNumber);


            //DateTime firstMonday = startDate;
            //while (firstMonday.DayOfWeek != DayOfWeek.Monday)
            //{
            //    firstMonday = firstMonday.AddDays(1);
            //}

            DateTime currentDate = new DateTime(2025, 1, 1);
            DateTime startDate = currentDate;
            while (startDate.ToString("ddd") != "Mon")
            {
                startDate = startDate.AddDays(1);
            }
           
            try
            {
                int LastYear = currentDate.AddYears(-1).Year;

                string folderPath = @"C:\Users\Public\Documents\Weekly SOS";

                // Get all files in the folder that contain "South" in their names
                List<string> files = Directory.GetFiles(folderPath)
                                              .Where(file => file.Contains("South"))
                                              .ToList();



                foreach (var file in files)
                {
                    // Open the Excel file
                    FileInfo existingFile = new FileInfo(file);

                    using (ExcelPackage package = new ExcelPackage(existingFile))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets["Weekly Summary"];

                        int lastUsedRow = worksheet.Dimension.End.Row;
                        int lastUsedColumn = worksheet.Dimension.End.Column;

                        int regionRow = -1;
                        for (int row = 1; row <= lastUsedRow; row++)
                        {
                            var regionValue = worksheet.Cells[row, 1].Value;
                            if (worksheet.Cells[row, 1].Text.Contains("Region"))
                            {
                                regionRow = row;
                                break;
                            }
                        }

                        int lastYearColumn = -1;
                        for (int col = 1; col <= lastUsedColumn; col++)
                        {
                            if (worksheet.Cells[regionRow, col].Text.Contains(LastYear.ToString()))
                            {
                                lastYearColumn = col;
                                break;
                            }
                        }

                        for (int col = 1; col <= lastUsedColumn; col++)
                        {
                            if (worksheet.Column(col).Hidden)
                            {
                                worksheet.Column(col).Hidden = false;
                            }
                        }

                        int startRow = 0;
                        for (int row = 1; row <= lastUsedRow; row++)
                        {
                            var values = worksheet.Cells[row, 4].Value;
                            if (values == null)
                            {
                                continue;
                            }
                            if (values is double)
                            {
                                startRow = row;
                                break;
                            }
                        }

                        int lastFormulaRow = 0;
                        for (int row = 1; row <= lastUsedRow; row++)
                        {
                            if (worksheet.Cells[1, 1].Text.Contains("Company Avg"))
                            {
                                lastFormulaRow = row;
                                break;
                            }
                        }

                        int insertColumn = lastYearColumn + 1;
                        worksheet.InsertColumn(insertColumn, 1);
                        var insertColumnName = GetExcelColumnName(insertColumn);

                        for (int row = 1; row <= lastUsedRow; row++)
                        {
                            worksheet.Cells[row, insertColumn].StyleID = worksheet.Cells[row, insertColumn - 1].StyleID;
                            if (!string.IsNullOrEmpty(worksheet.Cells[row, insertColumn -1].Formula))
                            {
                                if(worksheet.Cells[row, 1].Text.StartsWith("D"))
                                {
                                    worksheet.Cells[row, insertColumn].Formula = $"=IFERROR(AVERAGEIF($B${startRow}:$B${lastFormulaRow - 2},$B{row},{insertColumnName}${startRow}:{insertColumnName}${lastFormulaRow - 2}),\"-\")";
                                }
                                else
                                {
                                    worksheet.Cells[row, insertColumn].Formula = $"=IFERROR(AVERAGEIF($A${startRow}:$A${lastFormulaRow - 2},$A{row},{insertColumnName}${startRow}:{insertColumnName}${lastFormulaRow - 2}),\"-\")";
                                }
                            }

                            if(row == lastFormulaRow)
                            {
                                worksheet.Cells[row, insertColumn].Formula = $"=AVERAGE({insertColumnName}{startRow}:{insertColumnName}{lastFormulaRow-2})";
                            }
                        }

                        for (int row = 1; row <= lastUsedRow; row++)
                        {
                            worksheet.Cells[row, insertColumn - 1].Value = worksheet.Cells[row, insertColumn - 1].Value;
                            worksheet.Cells[row, insertColumn - 1].Value = worksheet.Cells[row, insertColumn - 1].Text;
                        }


                        for (int row = startRow; row <= lastUsedRow; row++)
                        {
                            for (int col = insertColumn + 2; col <= worksheet.Dimension.End.Column; col++)
                            {
                                worksheet.Cells[row, col].Value = null;
                                worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None;
                            }
                        }

                        int dateColumn = 0;
                        for (int col = 1; col <= lastUsedColumn; col++)
                        {
                            if(worksheet.Cells[regionRow - 1, col].Text == "1")
                            {
                                dateColumn = col;
                                break;
                            }
                        }

                        DateTime currentMonday = startDate;
                        while (currentMonday.Year == startDate.Year && dateColumn >= 1)
                        {
                            worksheet.Cells[regionRow, dateColumn].Value = currentMonday.ToString("MM/dd/yyyy");
                            currentMonday = currentMonday.AddDays(7); 
                            dateColumn--;
                        }

                        if (currentMonday.Year == 2026)
                        {
                            worksheet.InsertColumn(dateColumn + 1, 1);
                            var insert53ColumnName = GetExcelColumnName(dateColumn + 1);

                            for (int row = 1; row <= lastUsedRow; row++)
                            {
                                worksheet.Cells[row, dateColumn + 1].StyleID = worksheet.Cells[row, dateColumn + 2].StyleID;
                                if (!string.IsNullOrEmpty(worksheet.Cells[row, dateColumn + 2].Formula))
                                {
                                    if(worksheet.Cells[row, 1].Text.StartsWith("D"))
                                    {
                                        worksheet.Cells[row, dateColumn + 1].Formula = $"=IFERROR(AVERAGEIF($A${startRow}:$A${lastFormulaRow - 2},$A{row}, {insert53ColumnName}${startRow}:{insert53ColumnName}${lastFormulaRow - 2}), \"-\")";

                                    }
                                    else
                                    {
                                        worksheet.Cells[row, dateColumn + 1].Formula = $"=IFERROR(AVERAGEIF($B${startRow}:$B${lastFormulaRow - 2},$B{row}, {insert53ColumnName}${startRow}:{insert53ColumnName}${lastFormulaRow - 2}), \"-\")";
                                    }

                                }
                            }

                            worksheet.Cells[regionRow, dateColumn + 1].Value = currentMonday.ToString("MM/dd/yyyy");
                        }

                        string fileName = Path.GetFileName(file);

                        // Save the changes
                        package.SaveAs($@"C:\Users\Nimap\Documents\WeeklySOSYearlyTemplateUpdate\{fileName}");
                    }
                }

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static string GetExcelColumnName(int columnNumber)
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
}

        //static int GetYearOfWeekNumber(DateTime date, int weekNumber)
        //{
        //    Calendar cal = CultureInfo.CurrentCulture.Calendar;

//    // Determine the first day of the first week for the current year
//    DateTime firstDayOfYear = new DateTime(date.Year, 1, 1);
//    int firstWeekOfYear = cal.GetWeekOfYear(firstDayOfYear, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

//    // Determine the first day of the first week for the next year
//    DateTime firstDayOfNextYear = new DateTime(date.Year + 1, 1, 1);
//    int firstWeekOfNextYear = cal.GetWeekOfYear(firstDayOfNextYear, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

//    // If the current date's week number is less than the first week's week number of the next year, it's the current year
//    // Otherwise, it's the next year
//    if (weekNumber >= firstWeekOfNextYear)
//    {
//        return date.Year + 1;
//    }
//    else
//    {
//        return date.Year;
//    }
//}

