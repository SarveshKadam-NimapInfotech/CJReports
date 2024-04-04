using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace SalesTax
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.SalesTax();

        }

        public void SalesTax()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string salesTaxFilePath = @"";

            string glFilePath = @"";

            Excel.Workbook salesTaxWorkbook = excelApp.Workbooks.Open(salesTaxFilePath);

            Excel.Workbook glWorkbook = excelApp.Workbooks.Open(glFilePath);

            try
            {
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
