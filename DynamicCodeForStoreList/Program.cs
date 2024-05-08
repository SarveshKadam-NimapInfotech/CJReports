using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Xml.Linq;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace DynamicCodeForStoreList
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string storeListFilePath = @"C:\Users\Public\Documents\StoreList.xlsx";

            //using (var storeExcel = new ExcelPackage(new FileInfo(storeListFilePath)))
            //{
            //    var storeSheet = storeExcel.Workbook.Worksheets[0];
            //    //var presaleStoreSheet = presaleExcel.Workbook.Worksheets.Add("CJ Site List", storeSheet);

            //    // Find the index of "Region 1" (case insensitive)
            //    int rowIndex = 1; // Start from the first row
            //    int regionRowIndex = -1; // Initialize as -1 to indicate not found
            //    while (rowIndex <= storeSheet.Dimension.End.Row)
            //    {
            //        if (storeSheet.Cells[rowIndex, 1].Text.ToLower() == "region 1")
            //        {
            //            regionRowIndex = rowIndex;
            //            break;
            //        }
            //        rowIndex++;
            //    }

            //    if (regionRowIndex != -1)
            //    {
            //        int rowsToInsert = 8 - regionRowIndex; // Calculate the number of rows to insert

            //        if (rowsToInsert > 0)
            //        {
            //            // Insert rows above "Region 1"
            //            storeSheet.InsertRow(regionRowIndex, rowsToInsert);
            //        }
            //        else if (rowsToInsert < 0)
            //        {
            //            // Move "Region 1" to the 8th row
            //            storeSheet.DeleteRow(1, regionRowIndex - 8);
            //        }
            //        // If rowsToInsert == 0, "Region 1" is already on the 8th row, do nothing
            //    }
            //    else
            //    {
            //        // "Region 1" not found
            //        // Handle the case if necessary
            //    }

            //    storeExcel.SaveAs(@"C:\Users\Public\Documents\StoreListTest.xlsx");
            //}

            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            Workbook storeListWorkbook = excelApp.Workbooks.Open(storeListFilePath);
            Worksheet cjListing = storeListWorkbook.Worksheets[1];

            Dictionary<string, List<string>> dict = new Dictionary<string, List<string>>();

            int row = 6;
            while (cjListing.Cells[row, 1].Value != null)
            {
                string cellValue = Convert.ToString(cjListing.Cells[row, 1].Value);
                if (cellValue == "North")
                {
                    break;
                }

                if (cellValue.Contains("Region"))
                {
                    string key = cellValue;
                    key = "R" + cellValue.Substring(7);


                    List<string> values = new List<string>();

                    while (cjListing.Cells[++row, 1].Value != null)
                    {
                        string nextCellValue = Convert.ToString(cjListing.Cells[row, 1].Value);

                        if (nextCellValue.StartsWith("Region") || nextCellValue == "North")
                        {
                            break;
                        }
                        if (nextCellValue.StartsWith("Dist"))
                        {
                            string districtNumber = nextCellValue.Substring(4).Trim();

                            if (int.TryParse(districtNumber, out int _))
                            {
                                values.Add("D" + districtNumber);
                            }
                            
                        }
                        if (nextCellValue.StartsWith("D") && !nextCellValue.StartsWith("Dist"))
                        {
                            values.Add(nextCellValue);
                        }


                    }

                    dict.Add(key, values);
                }
            }


        }
    }
}
