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

            string salesTaxFilePath = @"C:\Users\Nimap\Documents\SalesTax\CA Sales Tax 1Qtr 2024-02 Excl EBT Sales.xlsx";

            string glFilePath = @"";

            Excel.Workbook salesTaxWorkbook = excelApp.Workbooks.Open(salesTaxFilePath);

            Excel.Workbook glWorkbook = excelApp.Workbooks.Open(glFilePath);

            try
            {
                var date = "02/29/2023";

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


                Worksheet salesTaxSummarySheet = salesTaxWorkbook.Worksheets["Summary"];

                salesTaxSummarySheet.Range["B2"].Value = month;
                salesTaxSummarySheet.Range["B3"].Value = year;
                salesTaxSummarySheet.Range["B3"].Value = date;

                switch (month)
                {
                    case 2:
                        if (previousMonth == 1)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"C8:C14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"D8:D14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 3:
                        if (previousMonth == 2)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"D8:C14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"E8:E14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 4:
                        if (previousMonth == 3)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"E8:E14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"F8:F14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 5:
                        if (previousMonth == 4)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"F8:F14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"G8:G14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 6:
                        if (previousMonth == 5)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"G8:G14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"H8:H14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 7:
                        if (previousMonth == 6)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"H8:H14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"I8:I14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 8:
                        if (previousMonth == 7)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"I8:I14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"J8:J14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 9:
                        if (previousMonth == 8)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"J8:J14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"K8:K14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 10:
                        if (previousMonth == 9)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"K8:K14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"L8:L14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 11:
                        if (previousMonth == 10)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"L8:L14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"M8:M14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    case 12:
                        if (previousMonth == 11)
                        {
                            Range copyRange1 = salesTaxSummarySheet.Range[$"M8:M14"];
                            Range pasteRange1 = salesTaxSummarySheet.Range[$"N8:N14"];

                            copyRange1.Copy(Type.Missing);
                            pasteRange1.PasteSpecial(XlPasteType.xlPasteFormulas);
                            copyRange1.PasteSpecial(XlPasteType.xlPasteValues);

                        }
                        break;

                    default:
                        break;
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
