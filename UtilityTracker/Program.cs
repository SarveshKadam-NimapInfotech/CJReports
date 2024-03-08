using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Media;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Excel = Microsoft.Office.Interop.Excel;

namespace UtilityTracker
{
    public class YourEntityClass
    {
        public string AccountNumber { get; set; }
        public List<EntityDataRow> DataRows { get; set; } = new List<EntityDataRow>();
    }

    public class EntityDataRow
    {
        public string Entity { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public DateTime BillDate { get; set; }
        public decimal Cost { get; set; }
        public int Usage { get; set; }
        public string Vendor { get; set; }

        public string Utility { get; set; }
    }


    internal class Program
    {
        static void Main(string[] args)
        {
            Program program = new Program();
            program.UtilityTracker();

            //program.UtilityTrackerEplus();
        
        }

        public void UtilityTracker()
        {
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Interactive = false;
            excelApp.DisplayAlerts = false;
            excelApp.DisplayClipboardWindow = false;
            excelApp.DisplayStatusBar = false;

            string costUsageFilePath = @"C:\Users\Nimap\Documents\UtilityTracker\InputFiles\Cost Usage.xlsx";

            string utilityTrackerFilePath = @"C:\Users\Nimap\Documents\UtilityTracker\InputFiles\RE - Utility Tracker (2023-09).xlsx";

            //string utilityTrackerFilePath = @"C:\Users\Nimap\Documents\UtilityTracker\RE - Utility Tracker (2023-09).xlsx";


            Excel.Workbook costUsageWorkbook = excelApp.Workbooks.Open(costUsageFilePath);

            Excel.Workbook utilityTrackerWorkbook = excelApp.Workbooks.Open(utilityTrackerFilePath);

            try
            {
                Worksheet utilitySheet = utilityTrackerWorkbook.Worksheets["Utility Data"];

                Range clearRange = utilitySheet.Range["A2:M" + utilitySheet.Rows.Count];
                clearRange.Clear();

                Worksheet trackerSheet = utilityTrackerWorkbook.Worksheets["Conservice tracking"];

                int trackerLastRow = trackerSheet.Cells[trackerSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 1;

                HashSet<string> uniqueAccountNumbers = new HashSet<string>();

                for (int i = 2; i < trackerLastRow; i++)
                {
                    string accountNo = Convert.ToString(trackerSheet.Cells[i, 3].Value);
                    uniqueAccountNumbers.Add(accountNo);
                }

                List<string> accountNumber = new List<string>(uniqueAccountNumbers);

                //foreach (string number in accountNumber)
                //{
                //    Console.WriteLine(number);
                //}

                Worksheet costUsageSheet = costUsageWorkbook.Worksheets[1];

                int costUsageLastRow = costUsageSheet.Cells[costUsageSheet.Rows.Count, 1].End[Excel.XlDirection.xlUp].Row + 1;

                List<YourEntityClass> matchedDataList = new List<YourEntityClass>();

                foreach (string accountNo in accountNumber)
                {
                    YourEntityClass entityData = new YourEntityClass
                    {
                        AccountNumber = accountNo,
                        DataRows = new List<EntityDataRow>()
                    };

                    for (int i = 2; i < costUsageLastRow; i++)
                    {
                        string dataAccountNo = Convert.ToString(costUsageSheet.Cells[i, 1].Value);

                        if (accountNo == dataAccountNo)
                        {
                            EntityDataRow rowData = new EntityDataRow
                            {
                                Entity = Convert.ToString(costUsageSheet.Cells[i, 4].Value),
                                StartDate = Convert.ToDateTime(costUsageSheet.Cells[i, 6].Value),
                                EndDate = Convert.ToDateTime(costUsageSheet.Cells[i, 7].Value),
                                BillDate = Convert.ToDateTime(costUsageSheet.Cells[i, 8].Value),
                                Cost = Convert.ToDecimal(costUsageSheet.Cells[i, 10].Value),
                                Usage = Convert.ToInt32(costUsageSheet.Cells[i, 11].Value),
                                Vendor = Convert.ToString(costUsageSheet.Cells[i, 13].Value),
                                Utility = Convert.ToString(costUsageSheet.Cells[i, 12].Value),

                            };

                            entityData.DataRows.Add(rowData);
                        }
                    }

                    if (entityData.DataRows.Count > 0)
                    {
                        matchedDataList.Add(entityData);
                    }
                }


                //foreach (YourEntityClass entityData in matchedDataList)
                //{
                //    Console.WriteLine($"Account Number: {entityData.AccountNumber}");

                //    foreach (EntityDataRow rowData in entityData.DataRows)
                //    {
                //        Console.WriteLine($"Entity: {rowData.Entity}");
                //        Console.WriteLine($"Start Date: {rowData.StartDate}");
                //        Console.WriteLine($"End Date: {rowData.EndDate}");
                //        Console.WriteLine($"Bill Date: {rowData.BillDate}");
                //        Console.WriteLine($"Cost: {rowData.Cost}");
                //        Console.WriteLine($"Usage: {rowData.Usage}");
                //        Console.WriteLine($"Vendor: {rowData.Vendor}");
                //        Console.WriteLine("--------------------------------");
                //    }

                //    Console.WriteLine("================================");
                //}

                //Worksheet utilitySheet = utilityTrackerWorkbook.Worksheets["Utility Data"];

                int row = 2;

                foreach (YourEntityClass entityData in matchedDataList)
                {
                    utilitySheet.Cells[row, 1].Value = entityData.AccountNumber;

                    foreach (EntityDataRow rowData in entityData.DataRows)
                    {
                        utilitySheet.Cells[row, 2].Value = rowData.Vendor;
                        utilitySheet.Cells[row, 4].Value = rowData.Entity;
                        utilitySheet.Cells[row, 6].Value = rowData.StartDate;
                        utilitySheet.Cells[row, 7].Value = rowData.EndDate;
                        utilitySheet.Cells[row, 5].Value = rowData.BillDate;
                        utilitySheet.Cells[row, 3].Value = rowData.Utility;
                        utilitySheet.Cells[row, 9].Value = rowData.Cost;
                        // utilitySheet.Cells[row, 9].Value = rowData.Usage;

                        int usage = rowData.Usage;

                        if (rowData.Utility.Contains("Water"))
                        {
                            usage /= 748;
                        }

                        utilitySheet.Cells[row, 10].Value = usage;

                        if (rowData.Utility.Contains("Electric"))
                        {
                            string nature = "Electricity";
                            utilitySheet.Cells[row, 11].Value = nature;

                        }

                        if (rowData.Utility.Contains("Gas"))
                        {
                            string nature = "Gas";
                            utilitySheet.Cells[row, 11].Value = nature;

                        }

                        if (rowData.Utility.Contains("Trash"))
                        {
                            string nature = "Trash";
                            utilitySheet.Cells[row, 11].Value = nature;

                        }

                        if (rowData.Utility.Contains("Water"))
                        {
                            string nature = "Water";
                            utilitySheet.Cells[row, 11].Value = nature;

                        }

                        if (rowData.Utility.Contains("Sewer"))
                        {
                            string nature = "Other";
                            utilitySheet.Cells[row, 11].Value = nature;

                        }


                        row++;
                    }

                    row++;
                }

                utilitySheet.Columns.AutoFit();

                //Worksheet utilitySheet = utilityTrackerWorkbook.Worksheets["Utility Data"];

                int utilityLastRow = utilitySheet.Cells[utilitySheet.Rows.Count, 2].End[Excel.XlDirection.xlUp].Row + 1;


                for (int i = 2; i < utilityLastRow; i++)
                {
                    string vendor = utilitySheet.Cells[i, 2].Value;
                    string entity = utilitySheet.Cells[i, 4].Value;
                    DateTime? startdate = utilitySheet.Cells[i, 6].Value;
                    DateTime? enddate = utilitySheet.Cells[i, 7].Value;
                    DateTime? billdate = utilitySheet.Cells[i, 5].Value;
                    string utility = utilitySheet.Cells[i, 3].Value;
                    decimal cost = Convert.ToDecimal(utilitySheet.Cells[i, 9].Value);
                    int usage = Convert.ToInt32(utilitySheet.Cells[i, 10].Value);

                    int checkRow = i;
                    while (vendor != null)
                    {
                        checkRow++;
                        string duplicateVendor = utilitySheet.Cells[checkRow, 2].Value;
                        string duplicateentity = utilitySheet.Cells[checkRow, 4].Value;
                        DateTime? duplicatestartdate = utilitySheet.Cells[checkRow, 6].Value;
                        DateTime? duplicateenddate = utilitySheet.Cells[checkRow, 7].Value;
                        DateTime? duplicatebilldate = utilitySheet.Cells[checkRow, 5].Value;
                        string duplicateutility = utilitySheet.Cells[i, 3].Value;
                        decimal duplicatecost = Convert.ToDecimal(utilitySheet.Cells[checkRow, 9].Value);
                        int duplicateusage = Convert.ToInt32(utilitySheet.Cells[checkRow, 10].Value);

                        if (duplicateVendor == null)
                        {

                            break;

                        }
                        if (vendor.Equals(duplicateVendor) && entity.Equals(duplicateentity) && startdate.Equals(duplicatestartdate) && enddate.Equals(duplicateenddate) && billdate.Equals(duplicatebilldate) && utility.Equals(duplicateentity))
                        {

                            cost = cost + duplicatecost;
                            usage = usage + duplicateusage;

                            utilitySheet.Cells[i, 9].Value = cost;
                            utilitySheet.Cells[i, 10].Value = usage;

                            utilitySheet.Rows[checkRow].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                            checkRow--;
                            utilityLastRow--;

                        }


                    }


                }

                Excel.Range helperColumnFormula = utilitySheet.Range[$"L2:L{utilityLastRow}"];
                helperColumnFormula.Formula = "=B2&K2&D2";




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

        public void UtilityTrackerEplus()
        {
            try
            {
                string costUsageFilePath = @"C:\Users\Nimap\Documents\UtilityTracker\InputFiles\Cost Usage.xlsx";

                //string utilityTrackerFilePath = @"C:\Users\Nimap\Documents\UtilityTracker\InputFiles\RE - Utility Tracker (2023-09).xlsx";

                string utilityTrackerFilePath = @"C:\Users\Nimap\Documents\UtilityTracker\RE - Utility Tracker (2023-09).xlsx";


                using (var utilityTrackerPackage = new ExcelPackage(new FileInfo(utilityTrackerFilePath)))
                using (var costUsagePackage = new ExcelPackage(new FileInfo(costUsageFilePath)))
                {
                    var utilitySheet = utilityTrackerPackage.Workbook.Worksheets["Utility Data"];
                    var trackerSheet = utilityTrackerPackage.Workbook.Worksheets["Conservice tracking"];
                    var costUsageSheet = costUsagePackage.Workbook.Worksheets[0];

                    utilitySheet.Cells[$"A2:I{utilitySheet.Dimension.End.Row}"].Clear();

                    int trackerLastRow = trackerSheet.Dimension.End.Row;

                    HashSet<string> uniqueAccountNumbers = new HashSet<string>();

                    for (int i = 2; i < trackerLastRow; i++)
                    {
                        string accountNo = Convert.ToString(trackerSheet.Cells[i, 3].Value);
                        uniqueAccountNumbers.Add(accountNo);
                    }

                    List<string> accountNumber = new List<string>(uniqueAccountNumbers);

                    int costUsageLastRow = costUsageSheet.Dimension.End.Row;

                    List<YourEntityClass> matchedDataList = new List<YourEntityClass>();

                    foreach (string accountNo in accountNumber)
                    {
                        YourEntityClass entityData = new YourEntityClass
                        {
                            AccountNumber = accountNo,
                            DataRows = new List<EntityDataRow>()
                        };

                        for (int i = 2; i < costUsageLastRow; i++)
                        {
                            string dataAccountNo = Convert.ToString(costUsageSheet.Cells[i, 1].Value);

                            if (accountNo == dataAccountNo)
                            {
                                EntityDataRow rowData = new EntityDataRow
                                {
                                    Entity = Convert.ToString(costUsageSheet.Cells[i, 4].Value),
                                    StartDate = Convert.ToDateTime(costUsageSheet.Cells[i, 6].Value),
                                    EndDate = Convert.ToDateTime(costUsageSheet.Cells[i, 7].Value),
                                    BillDate = Convert.ToDateTime(costUsageSheet.Cells[i, 8].Value),
                                    Cost = Convert.ToDecimal(costUsageSheet.Cells[i, 10].Value),
                                    Usage = Convert.ToInt32(costUsageSheet.Cells[i, 11].Value),
                                    Vendor = Convert.ToString(costUsageSheet.Cells[i, 13].Value),
                                    Utility = Convert.ToString(costUsageSheet.Cells[i, 12].Value),

                                };

                                entityData.DataRows.Add(rowData);
                            }
                        }

                        if (entityData.DataRows.Count > 0)
                        {
                            matchedDataList.Add(entityData);
                        }
                    }

                    int row = 2;

                    foreach (YourEntityClass entityData in matchedDataList)
                    {
                        utilitySheet.Cells[row, 1].Value = entityData.AccountNumber;

                        foreach (EntityDataRow rowData in entityData.DataRows)
                        {
                            utilitySheet.Cells[row, 2].Value = rowData.Vendor;
                            utilitySheet.Cells[row, 3].Value = rowData.Entity;
                            utilitySheet.Cells[row, 4].Value = rowData.StartDate;
                            utilitySheet.Cells[row, 5].Value = rowData.EndDate;
                            utilitySheet.Cells[row, 6].Value = rowData.BillDate;
                            utilitySheet.Cells[row, 7].Value = rowData.Utility;
                            utilitySheet.Cells[row, 8].Value = rowData.Cost;
                            utilitySheet.Cells[row, 9].Value = rowData.Usage;


                            row++;
                        }

                        row++;
                    }

                    utilitySheet.Cells[utilitySheet.Dimension.Address].AutoFitColumns();

                    int utilityLastRow = utilitySheet.Dimension.End.Row;

                    for (int i = 2; i < utilityLastRow; i++)
                    {
                        string vendor = utilitySheet.Cells[i, 2].Text;
                        string entity = utilitySheet.Cells[i, 3].Text;
                        DateTime? startdate = utilitySheet.Cells[i, 4].GetValue<DateTime?>();
                        DateTime? enddate = utilitySheet.Cells[i, 5].GetValue<DateTime?>();
                        DateTime? billdate = utilitySheet.Cells[i, 6].GetValue<DateTime?>();
                        string utility = utilitySheet.Cells[i, 7].Text;
                        decimal cost = Convert.ToDecimal(utilitySheet.Cells[i, 8].Text);
                        int usage = Convert.ToInt32(utilitySheet.Cells[i, 9].Text);

                        int checkRow = i;
                        while (vendor != null)
                        {
                            checkRow++;
                            string duplicateVendor = utilitySheet.Cells[checkRow, 2].Text;
                            string duplicateEntity = utilitySheet.Cells[checkRow, 3].Text;
                            DateTime? duplicateStartDate = utilitySheet.Cells[checkRow, 4].GetValue<DateTime?>();
                            DateTime? duplicateEndDate = utilitySheet.Cells[checkRow, 5].GetValue<DateTime?>();
                            DateTime? duplicateBillDate = utilitySheet.Cells[checkRow, 6].GetValue<DateTime?>();
                            string duplicateUtility = utilitySheet.Cells[checkRow, 7].Text;
                            decimal duplicateCost = Convert.ToDecimal(utilitySheet.Cells[checkRow, 8].Text);
                            int duplicateUsage = Convert.ToInt32(utilitySheet.Cells[checkRow, 9].Text);

                            if (duplicateVendor == null)
                            {

                                break;

                            }
                            if (vendor.Equals(duplicateVendor) && entity.Equals(duplicateEntity) && startdate.Equals(duplicateStartDate) && enddate.Equals(duplicateEndDate) && billdate.Equals(duplicateBillDate) && utility.Equals(duplicateUtility))
                            {

                                cost = cost + duplicateCost;
                                usage = usage + duplicateUsage;

                                utilitySheet.Cells[i, 7].Value = cost;
                                utilitySheet.Cells[i, 8].Value = usage;

                                utilitySheet.DeleteRow(checkRow, 1);
                                checkRow--;
                                utilityLastRow--;

                            }


                        }


                    }



                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message); ;
            }

        }
    }
}

