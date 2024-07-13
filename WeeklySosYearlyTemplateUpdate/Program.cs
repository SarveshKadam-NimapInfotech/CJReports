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
            string date = "12/30/2024";

            DateTime startDate;
            DateTime.TryParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out startDate);
            Calendar cal = CultureInfo.CurrentCulture.Calendar;
            int weekNumber = cal.GetWeekOfYear(startDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            // Determine the year of the given week number
            int year = GetYearOfWeekNumber(startDate, weekNumber);


            DateTime firstMonday = startDate;
            while (firstMonday.DayOfWeek != DayOfWeek.Monday)
            {
                firstMonday = firstMonday.AddDays(1);
            }

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
                    
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (worksheet.Column(col).Hidden)
                        {
                            worksheet.Column(col).Hidden = false;
                        }
                    }

                    worksheet.InsertColumn(6, 1);

                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        worksheet.Cells[row, 6].StyleID = worksheet.Cells[row, 5].StyleID; 
                        if (!string.IsNullOrEmpty(worksheet.Cells[row, 5].Formula))
                        {
                            worksheet.Cells[row, 6].Formula = worksheet.Cells[row, 5].Formula; 
                        }
                    }

                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                    {
                        worksheet.Cells[row, 5].Value = worksheet.Cells[row, 5].Value; 
                        worksheet.Cells[row, 5].Value = worksheet.Cells[row, 5].Text; 
                    }

                    for (int row = 17; row <= worksheet.Dimension.End.Row; row++)
                    {
                        for (int col = 7; col <= worksheet.Dimension.End.Column; col++) 
                        {
                            worksheet.Cells[row, col].Value = null; 
                            worksheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.None; 
                        }
                    }

                    // Add the dates of every Monday till the end of the year starting from the last used row in column E
                    DateTime currentMonday = firstMonday;
                    int dateColumn = worksheet.Dimension.End.Column; // Starting row for dates is the last used row in column E

                    while (currentMonday.Year == startDate.Year && dateColumn >= 1)
                    {
                        worksheet.Cells[5, dateColumn].Value = currentMonday.ToString("MM/dd/yyyy");
                        currentMonday = currentMonday.AddDays(7); // Move to the next Monday
                        dateColumn--; // Move to the previous column
                    }

                    string fileName = Path.GetFileName(file);

                    // Save the changes
                    package.SaveAs($@"C:\Users\Nimap\Documents\WeeklySOSYearlyTemplateUpdate\{fileName}");
                }
            }

        }

        static int GetYearOfWeekNumber(DateTime date, int weekNumber)
        {
            Calendar cal = CultureInfo.CurrentCulture.Calendar;

            // Determine the first day of the first week for the current year
            DateTime firstDayOfYear = new DateTime(date.Year, 1, 1);
            int firstWeekOfYear = cal.GetWeekOfYear(firstDayOfYear, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            // Determine the first day of the first week for the next year
            DateTime firstDayOfNextYear = new DateTime(date.Year + 1, 1, 1);
            int firstWeekOfNextYear = cal.GetWeekOfYear(firstDayOfNextYear, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            // If the current date's week number is less than the first week's week number of the next year, it's the current year
            // Otherwise, it's the next year
            if (weekNumber >= firstWeekOfNextYear)
            {
                return date.Year + 1;
            }
            else
            {
                return date.Year;
            }
        }
    }
}
