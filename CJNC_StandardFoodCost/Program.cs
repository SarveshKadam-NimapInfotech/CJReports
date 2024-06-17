using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CJNC_StandardFoodCost
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

            var templatePath = @"C:\Users\Nimap\Documents\CjncStandardFoodCost\previousMonthCjLabourStandardFoodCost-1713186755085.xlsx";

            var pnlFilePath = @"C:\Users\Nimap\Documents\CjncStandardFoodCost\pnlFile-1713186755547.xlsx";

            var glFilePath = @"C:\Users\Nimap\Documents\CjncStandardFoodCost\allLiveGl-1713186755584.xlsx";

            Excel.Workbook templateWorkbook = excelApp.Workbooks.Open(templatePath);
            Excel.Workbook pnlWorkbook = excelApp.Workbooks.Open(pnlFilePath);
            Excel.Workbook glWorkbook = excelApp.Workbooks.Open(glFilePath);

            try
            {
                string date = "03/29/2024";
                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                int monthInt = parsedDate.Month;

                string month = Convert.ToString(parsedDate.Month);

                string year = Convert.ToString(parsedDate.Year);

                int previousMonth = 0;
                if (monthInt == 1)
                {
                    previousMonth = 12;
                }
                else
                {
                    previousMonth = monthInt - 1;

                }

                string monthName = parsedDate.ToString("MMM", CultureInfo.InvariantCulture);


                Worksheet salesTaxSummarySheet = salesTaxWorkbook.Worksheets["Labor Standard Variance"];

                // Insert a new column at column J
                Excel.Range columnJ = salesTaxSummarySheet.Columns["J"];
                columnJ.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                Excel.Range columnK = salesTaxSummarySheet.Columns["K"];
                Excel.Range newColumnJ = salesTaxSummarySheet.Columns["J"];

                columnK.Copy(Type.Missing);
                newColumnJ.PasteSpecial(XlPasteType.xlPasteAll);
                copyRange1.PasteSpecial(XlPasteType.xlPasteValues);



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
