using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace FrmFlexBudget
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.FrmFlexBudget();
        }

        private void FrmFlexBudget()
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
