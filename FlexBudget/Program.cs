using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace FlexBudget
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.FlexBudget();
        }

        public void FlexBudget()
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
                var date = "12/01/2023";

                DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

                string month = parsedDate.Month.ToString();

                Worksheet RawDataPlSheet = flexBudgetWorkbook.Worksheets["Raw Data PL"];

                Worksheet CjNorthSheet = CjNorthWorkbook.Worksheets[1];

                Range clearRange = RawDataPlSheet.Range["B1:S" + RawDataPlSheet.Rows.Count];
                clearRange.Clear();

                Range copyRange = CjNorthSheet.Range["A1:R" + CjNorthSheet.Rows.Count];
                copyRange.Copy(Type.Missing);

                Range pasteRange = RawDataPlSheet.Range["B1:S" + RawDataPlSheet.Rows.Count];
                pasteRange.PasteSpecial(XlPasteType.xlPasteAll);

                Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

                int rawDataLastRow = RawDataPlSheet.Cells[RawDataPlSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 1;

                for (int i = 1; i <= rawDataLastRow; i++)
                {
                    var cellB = RawDataPlSheet.Cells[i, 2];
                    var cellC = RawDataPlSheet.Cells[i, 3];
                    var cellD = RawDataPlSheet.Cells[i, 4];

                    // Check if any of the cells are null
                    if (cellB.Value != null)
                    {
                        string cellValueB = Convert.ToString(cellB.Value);
                        string cellValueC = Convert.ToString(cellC?.Value);
                        string cellValueD = Convert.ToString(cellD?.Value);

                        if (cellValueB.Contains("Net Sales") || cellValueB.Contains("Labor Matrix") || cellValueB.Contains("Ideal Food Cost"))
                        {
                            string key = cellValueB;

                            if (!dict.ContainsKey(key))
                            {
                                dict[key] = new List<string>();
                            }

                            dict[key].Add(cellValueC ?? "");  // Value1
                            dict[key].Add(cellValueD ?? "");  // Value2
                        }
                    }
                }


                foreach (var kvp in dict)
                {
                    Console.WriteLine($"Key: {kvp.Key}");

                    List<string> values = kvp.Value;

                    for (int j = 0; j < values.Count; j += 2)
                    {
                        string value1 = values[j];
                        string value2 = (j + 1 < values.Count) ? values[j + 1] : "N/A";

                        Console.WriteLine($"  Value1: {value1}, Value2: {value2}");
                    }

                    Console.WriteLine();
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
