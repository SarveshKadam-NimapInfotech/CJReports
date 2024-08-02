using OfficeOpenXml;
using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Linq;
using OfficeOpenXml.FormulaParsing;

namespace Stock_Report
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string filePath = $@"C:\Users\Nimap\Documents\StockReport\Input files\2404_Stocks-Consolidated_-_04-30-2024_v11_rmUNfQU.xlsx";
            string targetPath = $@"C:\Users\Nimap\Documents\StockReport\Test Output\2404 Stocks-Consolidated - 04-30-2024 v8 Test.xlsx";
            string sourcePath = $@"C:\Users\Nimap\Documents\StockReport\Input files\Transactions_Saturday_June_01_2024_6-15PM_oIbGEfJ.xlsx";
            string date = "05/2024";

            bool isNewYear = false;
            DateTime currentDate = new DateTime(Convert.ToInt16(date.Split('/')[1]), Convert.ToInt16(date.Split('/')[0]), 1).AddMonths(1).AddDays(-1);
            DateTime previousMonth = currentDate.AddMonths(-1);
            DateTime nextMonth = currentDate.AddMonths(1);
            Dictionary<string, List<DateTime>> boughtStocksWithDate = new Dictionary<string, List<DateTime>>();
            //List<KeyValuePair<string, DateTime>> boughtStocksWithDate = new List<KeyValuePair<string, DateTime>>();
            Dictionary<string, List<double>> FinalValuesAfterSold = new Dictionary<string, List<double>>();
            Dictionary<string, double> SoldStocks = new Dictionary<string, double>();
            List<string> Stocks = new List<string>();
            List<string> Accounts = new List<string>();
            List<string> Quantities = new List<string>();
            List<string> Values = new List<string>();
            List<string> PurchaseValues = new List<string>();
            List<string> BoughtStocks = new List<string>();
            List<string> BoughtAccounts = new List<string>();
            List<string> BoughtQuantities = new List<string>();
            List<string> BoughtValues = new List<string>();

            //Extract the data from the lpl files for the consolidated row line items values.
            Dictionary<string, double> ConsolidateSheetLplValues = new Dictionary<string, double>();

            double transferAmount = 0;
            double dividendInterestAmount = 0;
            //THis Fee Amount is already in negative so we need to convert it first into positive and should add it.
            double FeeAmount = 0;
            double profitDivideLoss = 0;

            while (previousMonth.ToString("MMM") != currentDate.ToString("MMM"))
            {
                previousMonth = previousMonth.AddDays(1);
            }
            while (nextMonth.ToString("MMM") != currentDate.AddMonths(2).ToString("MMM"))
            {
                nextMonth = nextMonth.AddDays(1);
            }
            nextMonth = nextMonth.AddDays(-1);
            previousMonth = previousMonth.AddDays(-1);
            string formatDateCurrentMonth = $"{currentDate.ToString("MM")}/{currentDate.ToString("dd")}/{currentDate.ToString("yy")}";
            string formatDatePrevMonth = $"{previousMonth.ToString("MM")}/{previousMonth.ToString("dd")}/{previousMonth.ToString("yy")}";
            string formatDateNextMonth = $"{nextMonth.ToString("MM")}/{nextMonth.ToString("dd")}/{nextMonth.ToString("yy")}";
            string formatForStockHoldingHeader = $"{currentDate.ToString("MM")}/{currentDate.ToString("dd")}";
            Dictionary<string, int> ColumnMapper = new Dictionary<string, int>();
            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(filePath)))
                {

                    List<string> MainTableHeaders = new List<string>
                    {
                      "Sr.",
                      "Stock",
                       "Acct #",
                      $"No of Stock on {formatDatePrevMonth}",
                       $"Value as on {formatDatePrevMonth}",
                        "Buy $",
                            "Sale $",
                            "Transfer $",
                            "Profit/ Loss $",
                            "Mkt Price Per Unit $",
                            $"Value as on {formatDateCurrentMonth}",
                            "MTM Profit/(Loss) $",
                            "Chg %",
                            $"No of Stock on {formatDateCurrentMonth}",
                            "Original Value of Purchase",
                            "LTD Chg $",
                            "LTD Chg %"
                    };
                    List<string> SideTableHeaders = new List<string>
                    {
                        "Stock",
                        "Fund Invst",
                        "Units",
                        "Pur.Price Per Stock",
                        "Pur Date",
                        $"Sh Price on {formatForStockHoldingHeader}",
                        "LTD Profit/ Loss",
                        "Current Date",
                        "Days Holding",
                        "LTD %",
                        "YOY %"
                    };


                    ExcelWorkbook wb = excelPackage.Workbook;
                    ExcelWorksheet individualWs = wb.Worksheets["Individual Stocks"];
                    ExcelWorksheet stocksConsolidatedWs = wb.Worksheets["Stocks Consolidated"];

                    if (currentDate.Month == 1)
                    {
                        isNewYear = true;
                        GroupPreviousYearMonthsConsolidatedSheet(stocksConsolidatedWs, currentDate);
                    }
                    //Calculate the Worksheet as needed all the values
                    individualWs.Calculate();
                    string findMonth = $"{previousMonth.ToString("MMM")}'{previousMonth.ToString("yy")}";

                    bool isPrevMonth = false;
                    for (int row = 1; row <= individualWs.Dimension.End.Row; row++)
                    {
                        var value = individualWs.Cells[row, 2].Value;
                        if (value == null)
                        {
                            continue;
                        }
                        if (value.ToString().Contains(findMonth))
                        {
                            isPrevMonth = true;
                            row++;
                            continue;
                        }

                        if (!isPrevMonth)
                        {
                            continue;
                        }
                        if (value.ToString().ToLower().Equals("stock"))
                        {
                            ColumnMapper = IndividualStocksColumnsMapper(row, individualWs, ColumnMapper, formatDatePrevMonth);
                            continue;
                        }
                        if (value.ToString().ToLower().Contains("total (unrealised)"))
                        {
                            break;
                        }
                        Stocks.Add(value.ToString());
                        Accounts.Add(individualWs.Cells[row, 3].Value.ToString());
                        Values.Add(individualWs.Cells[row, ColumnMapper["valueCol"]].Value.ToString());
                        Quantities.Add(individualWs.Cells[row, ColumnMapper["stockNoCol"]].Value.ToString());
                        PurchaseValues.Add(individualWs.Cells[row, ColumnMapper["originalPurchase"]].Value.ToString());

                    }

                    using (ExcelPackage sourcePackage = new ExcelPackage(new FileInfo(sourcePath)))
                    {
                        ExcelWorkbook sourceWb = sourcePackage.Workbook;
                        ExcelWorksheet sourceWs = sourceWb.Worksheets[0];
                        //Filterting Data are as follow Funds, Buy, Redemeption, sell,interest,cash dividend.
                        for (int row = 1; row <= sourceWs.Dimension.End.Row; row++)
                        {
                            var activityCell = sourceWs.Cells[row, 6].Value;
                            var accountCell = sourceWs.Cells[row, 2].Value;
                            if (accountCell == null || activityCell == null)
                            {
                                continue;
                            }
                            accountCell = AccountNumberSeperationByDash(accountCell.ToString());

                            if (activityCell.ToString().ToLower().Contains("buy"))
                            {
                                string name = IdentifyName(sourceWs.Cells[row, 1].Value.ToString());
                                string accountWithName = $"{sourceWs.Cells[row, 3].Value.ToString()}-{accountCell.ToString()}-{name}";
                                DateTime buyDate = ConversionOfValuesToDate(sourceWs.Cells[row, 4].Value);
                                string stockName = $"{sourceWs.Cells[row, 7].Value.ToString()} {sourceWs.Cells[row, 8].Value.ToString().Trim()}";
                                string amount = Convert.ToString(-Math.Round(Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000));
                                string quantity = sourceWs.Cells[row, 10].Value.ToString();
                                string compositeKey = $"Stock:{stockName}Account:{accountWithName}"; 
                                BoughtStocks.Add(stockName);
                                BoughtAccounts.Add(accountWithName);
                                BoughtValues.Add(amount);
                                BoughtQuantities.Add(quantity);
                                if (!boughtStocksWithDate.ContainsKey(compositeKey))
                                {
                                    boughtStocksWithDate.Add(compositeKey, new List<DateTime>());
                                }
                                boughtStocksWithDate[compositeKey].Add(buyDate);
                                //boughtStocksWithDate.Add(compositeKey, buyDate);
                                //boughtStocksWithDate.Add(new KeyValuePair<string, DateTime>(compositeKey, buyDate));

                            }
                            else if (activityCell.ToString().ToLower().StartsWith("ach"))
                            {
                                transferAmount += Math.Round(Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000);
                                if (!ConsolidateSheetLplValues.ContainsKey("TransferOut"))
                                {
                                    ConsolidateSheetLplValues.Add("TransferOut", Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000);
                                }
                                else
                                {

                                    ConsolidateSheetLplValues["TransferOut"] += Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000;
                                }
                            }
                            else if (activityCell.ToString().ToLower().Trim().StartsWith("interest"))
                            {
                                dividendInterestAmount += Convert.ToDouble(sourceWs.Cells[row, 12].Value);
                                if (!ConsolidateSheetLplValues.ContainsKey("Interest"))
                                {
                                    ConsolidateSheetLplValues.Add("Interest", Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000);
                                }
                                else
                                {

                                    ConsolidateSheetLplValues["Interest"] += Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000;
                                }

                            }
                            else if (activityCell.ToString().ToLower().Trim().StartsWith("tax")) // Tax Condition added for Wire Fees
                            {
                                dividendInterestAmount += Convert.ToDouble(sourceWs.Cells[row, 12].Value);
                                if (!ConsolidateSheetLplValues.ContainsKey("Tax"))
                                {
                                    ConsolidateSheetLplValues.Add("Tax", Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000);
                                }
                                else
                                {

                                    ConsolidateSheetLplValues["Tax"] += Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000;
                                }

                            }
                            else if (activityCell.ToString().ToLower().Trim().Contains("dividend"))
                            {
                                dividendInterestAmount += Convert.ToDouble(sourceWs.Cells[row, 12].Value);
                                if (!ConsolidateSheetLplValues.ContainsKey("Dividend"))
                                {
                                    ConsolidateSheetLplValues.Add("Dividend", Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000);
                                }
                                else
                                {

                                    ConsolidateSheetLplValues["Dividend"] += Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000;
                                }

                            }
                            else if (activityCell.ToString().ToLower().Trim().StartsWith("fee"))
                            {
                                //THis Fee Amount is already in negative so we need to convert it first into positive and should add it.
                                FeeAmount += -Convert.ToDouble(sourceWs.Cells[row, 12].Value);
                                if (!ConsolidateSheetLplValues.ContainsKey("FeesExpenses"))
                                {
                                    ConsolidateSheetLplValues.Add("FeesExpenses", Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000);
                                }
                                else
                                {

                                    ConsolidateSheetLplValues["FeesExpenses"] += Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000;
                                }
                            }
                            else if (activityCell.ToString().ToLower().StartsWith("sell") || activityCell.ToString().ToLower().StartsWith("redemption"))
                            {
                                //Checking for the stock in the individual sheet.
                                string stockName = $"{sourceWs.Cells[row, 7].Value.ToString()} {sourceWs.Cells[row, 8].Value.ToString().Trim()}";


                                double sellAmount = Math.Round(Convert.ToDouble(sourceWs.Cells[row, 12].Value) / 1000);
                                double quantity = -Convert.ToDouble(sourceWs.Cells[row, 10].Value);
                                string nickName = IdentifyName(sourceWs.Cells[row, 1].Value.ToString());
                                string accountNo = $"{sourceWs.Cells[row, 3].Value.ToString()}-{accountCell.ToString()}-{nickName}";
                                //Check if the the stock name and account number is matching then go further for changing the quantity and amount.
                                if (Stocks.Contains(stockName) && Accounts.Contains(accountNo))
                                {
                                    int index = Stocks.FindIndex(stock => stock == stockName);
                                    double stockAmountPrevMonth = Convert.ToDouble(Values[index]);
                                    double stockQuantity = Convert.ToDouble(Quantities[index]);
                                    double purchaseValue = Convert.ToDouble(PurchaseValues[index]);
                                    //Finding the stock in the individual sheet Stocks of the previous month.
                                    if (index != -1)
                                    {
                                        //If the stock quantity matches the lpl file quantity then we will directly remove that stock and add in the sold stocks Dictionary.
                                        if (stockQuantity == quantity)
                                        {
                                            string accountHolder = $"{sourceWs.Cells[row, 3].Value.ToString()}-{accountCell.ToString()}-{nickName}";
                                            string compositeValues = $"Stock:{stockName}Account:{accountHolder}OriginalPurchase:{purchaseValue}StockPrevValues:{stockAmountPrevMonth}Quantity:{quantity}";
                                            SoldStocks.Add(compositeValues, sellAmount);
                                            Stocks.RemoveAt(index);
                                            Quantities.RemoveAt(index);
                                            Values.RemoveAt(index);
                                            Accounts.RemoveAt(index);
                                            PurchaseValues.RemoveAt(index);
                                        }
                                        else if (quantity > stockQuantity)
                                        {
                                            double totalQuantities = 0;
                                            double tempQuantity = quantity;
                                            double purchaseAmount = 0;
                                            double prevMonthAmount = 0;
                                            //If the quantity is greater than the actual stock quantity then we will use while loop and remove all the stock whose quantity is then the lpl file stock quantity
                                            while (totalQuantities < tempQuantity)
                                            {
                                                int secondIndex = Stocks.FindIndex(stock => stock == stockName);
                                                double minQuantity = Convert.ToDouble(Quantities[secondIndex]);
                                                double minAmount = Convert.ToDouble(Values[secondIndex]);
                                                double minPurchaseAmount = Convert.ToDouble(PurchaseValues[secondIndex]);
                                                string minStock = Stocks[secondIndex];
                                                string minAccount = Accounts[secondIndex];
                                                if (secondIndex != -1)
                                                {
                                                    if (minQuantity < tempQuantity)
                                                    {
                                                        Stocks.RemoveAt(secondIndex);
                                                        Quantities.RemoveAt(secondIndex);
                                                        Values.RemoveAt(secondIndex);
                                                        Accounts.RemoveAt(secondIndex);
                                                        PurchaseValues.RemoveAt(secondIndex);
                                                        tempQuantity -= minQuantity;
                                                        purchaseAmount = minPurchaseAmount;
                                                        prevMonthAmount = minAmount;
                                                    }
                                                    else
                                                    {
                                                        double varianceValue = minQuantity - tempQuantity;
                                                        double amountToMinus = varianceValue * minAmount / minAmount;
                                                        minAmount = -amountToMinus;
                                                        //FinalValuesAfterSold.Add($"Stock:{minStock}Account:{minAccount}Amount:{minAmount}", varianceValue);
                                                        FinalValuesAfterSold.Add($"Stock:{minStock}Account:{minAccount}", new List<double>());
                                                        //At 0 Indexed we have amount value;
                                                        FinalValuesAfterSold[$"Stock:{minStock}Account:{minAccount}"].Add(minAmount);
                                                        //At 1 Indexed we have Quantity Value;
                                                        FinalValuesAfterSold[$"Stock:{minStock}Account:{minAccount}"].Add(varianceValue);
                                                        purchaseAmount = minPurchaseAmount;
                                                        prevMonthAmount = minAmount;

                                                    }
                                                    totalQuantities += minQuantity;
                                                }
                                            }
                                            string accountHolder = $"{sourceWs.Cells[row, 3].Value.ToString()}-{accountCell.ToString()}-{nickName}";

                                            string compositeValues = $"Stock:{stockName}Account:{accountHolder}OriginalPurchase:{purchaseAmount}StockPrevValues:{prevMonthAmount}Quantity:{quantity}";
                                            SoldStocks.Add(compositeValues, sellAmount);
                                        }
                                    }

                                }

                            }
                        }
                    }
                    //Getting the cash row of the previous month and calculating the table also of the previous month in the individual sheet.
                    int cashRowPrevMonth = GetCashRowOfPrevTable(individualWs);
                    //Get the Value as on for the current month cash value .
                    double currentMonthCashValue = Convert.ToDouble(individualWs.Cells[cashRowPrevMonth, 11].Value);

                    //Convert all the fee and dividend,interet row values into profit and loss for profit and loss value of the cash row in the current month
                    profitDivideLoss = (dividendInterestAmount - FeeAmount) / 1000;

                    int monthHeaderRows = 3;
                    int totalHeaderLineItem = 1;
                    int cashRow = 1;
                    int totalSpacesRows = 2;
                    int TotalRowToAdd = monthHeaderRows + Stocks.Count + BoughtStocks.Count + totalHeaderLineItem + SoldStocks.Count + totalHeaderLineItem + cashRow + totalHeaderLineItem + totalSpacesRows;
                    individualWs.InsertRow(1, TotalRowToAdd);
                    int headersRow = 3;
                    int insertingRow = 4;
                    int srNo = 1;
                    string currentMonthString = $"{currentDate.ToString("MMM")}'{currentDate.ToString("yy")}";
                    individualWs.Cells[1, 2].Value = currentMonthString;
                    TitleFormatter(individualWs.Cells[1, 2]);
                    individualWs.Cells[2, 2].Value = "(Amt '000)";
                    SecondLineTitleFormatter(individualWs.Cells[2, 2]);
                    int startingCol = 1;
                    for (int i = 0; i < MainTableHeaders.Count; i++)
                    {
                        string data = MainTableHeaders[i];
                        individualWs.Cells[headersRow, startingCol].Value = data;
                        TableColumnHeadersFormatter(individualWs.Cells[headersRow, startingCol]);
                        startingCol++;
                    }

                    startingCol += 2;
                    int sideTableCol = startingCol;
                    for (int i = 0; i < SideTableHeaders.Count; i++)
                    {
                        string data = SideTableHeaders[i];
                        individualWs.Cells[headersRow, startingCol].Value = data;
                        TableColumnHeadersFormatter(individualWs.Cells[headersRow, startingCol]);
                        startingCol++;
                    }
                    var PreviousMonthStockTuple = Tuple.Create(Stocks, Accounts, Quantities, Values, PurchaseValues);
                    var CurrentMonthBoughtStockTuple = Tuple.Create(BoughtStocks, BoughtAccounts, BoughtQuantities, BoughtValues);
                    var cashRowValuesTuple = Tuple.Create(currentMonthCashValue, transferAmount, profitDivideLoss);
                    var MainTableTuple = await CreateMainTable(individualWs, insertingRow, srNo, PreviousMonthStockTuple, CurrentMonthBoughtStockTuple, cashRowValuesTuple, SoldStocks, FinalValuesAfterSold);
                    var SideTableTuple = Tuple.Create(Stocks, Accounts);
                    CreateSideTable(individualWs, sideTableCol, insertingRow, currentDate, SideTableTuple, BoughtStocks, BoughtAccounts, boughtStocksWithDate);
                    CreateConsolidatedSheet(stocksConsolidatedWs, currentDate, MainTableTuple, ConsolidateSheetLplValues, isNewYear);
                    excelPackage.SaveAs(targetPath);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        public static string AccountNumberSeperationByDash(string accountNumber)
        {
            string result = string.Empty;
            for (int i = 0; i < accountNumber.Length; i++)
            {
                if (i == accountNumber.Length / 2)
                {
                    result += "-";
                    result += accountNumber[i];
                }
                else
                {
                    result += accountNumber[i];
                }
            }
            return result;
        }

        public static Dictionary<string, int> IndividualStocksColumnsMapper(int rowNumber, ExcelWorksheet ws, Dictionary<string, int> ColumnMapper, string formatDate)
        {
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var values = ws.Cells[rowNumber, col].Value;
                if (values == null)
                {
                    continue;
                }
                if (values.ToString().ToLower().Equals($"value as on {formatDate}"))
                {
                    ColumnMapper.Add("valueCol", col);
                }
                else if (values.ToString().ToLower().Equals($"no of stock on {formatDate}"))
                {
                    ColumnMapper.Add("stockNoCol", col);
                }
                else if (values.ToString().ToLower().StartsWith("original"))
                {
                    ColumnMapper.Add("originalPurchase", col);
                }
            }
            return ColumnMapper;
        }

        public static string IdentifyName(string value)
        {
            string nickName = string.Empty;
            if (value.ToLower().Contains("neha"))
            {
                nickName = "NEHA";
            }
            else if (value.ToLower().Contains("zaver"))
            {
                nickName = "ZAVER";
            }
            else
            {
                nickName = "DFL";
            }

            return nickName;
        }

        public static DateTime ConversionOfValuesToDate(object value)
        {
            DateTime date = default;
            if (value is double valueDouble)
            {
                date = DateTime.FromOADate(valueDouble).Date;
            }
            else if (value is string valueString)
            {
                if (DateTime.TryParse(valueString, out DateTime valueDate))
                {
                    date = valueDate;
                }
            }
            else if (value is DateTime valueDate)
            {
                date = valueDate.Date;
            }
            return date;
        }

        public static async Task<double> Scrapping(string symbol)
        {
            double value = 0;
            var url = $"https://finance.yahoo.com/quote/{symbol}/";
            var handler = new HttpClientHandler
            {
                AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate
            };
            // Create an HttpClient to fetch the page content
            using (var httpClient = new HttpClient(handler))
            {
                var response = await httpClient.GetStringAsync(url);

                // Load the content into HtmlDocument
                var htmlDoc = new HtmlDocument();
                htmlDoc.LoadHtml(response);

                // Select the node containing the stock price value using the attributes you provided
                var priceNode = htmlDoc.DocumentNode.SelectSingleNode("//fin-streamer[@data-field='regularMarketPrice']");

                if (priceNode != null)
                {
                    // Get the value from the node's attribute 'data-value'
                    var priceValue = priceNode.GetAttributeValue("data-value", "not found");
                    value = Convert.ToDouble(priceValue);
                }
            }
            return value;
        }
        public static async Task<Tuple<Dictionary<string, List<int>>, Dictionary<string, double>>> CreateMainTable(OfficeOpenXml.ExcelWorksheet ws, int row, int srNo, Tuple<List<string>, List<string>, List<string>, List<string>, List<string>> StockData, Tuple<List<string>, List<string>, List<string>, List<string>> BoughtStockData, Tuple<double, double, double> CashRowValues, Dictionary<string, double> SoldStocks, Dictionary<string, List<double>> FinalValuesAfterSoldStocks)
        {
            List<string> Stocks = StockData.Item1;
            List<string> Accounts = StockData.Item2;
            List<string> Quantities = StockData.Item3;
            List<string> Values = StockData.Item4;
            List<string> PurchaseValues = StockData.Item5;

            List<string> BoughtStocks = BoughtStockData.Item1;
            List<string> BoughtAccounts = BoughtStockData.Item2;
            List<string> BoughtQuantities = BoughtStockData.Item3;
            List<string> BoughtValues = BoughtStockData.Item4;

            double cashValueCurrentMonth = CashRowValues.Item1;
            double transferValue = CashRowValues.Item2;
            double profitAndLossValue = CashRowValues.Item3;

            Dictionary<string, List<int>> ConsolidateSheetFormulaValues = new Dictionary<string, List<int>>();

            Dictionary<string, double> ConsolidateSheetValues = new Dictionary<string, double>();

            for (int i = 0; i < Stocks.Count; i++)
            {
                bool isUnited = false;
                bool isUSLineItems = false;
                bool isUnitedZaver = false;

                if (Stocks[i].ToLower().Trim().EndsWith("united") && Accounts[i].ToLower().Trim().EndsWith("dfl"))
                {
                    isUnited = true;
                }
                if (Stocks[i].ToLower().Trim().Contains("u s") && Stocks[i].ToLower().Trim().StartsWith("9"))
                {
                    isUSLineItems = true;
                }
                if (Stocks[i].ToLower().Trim().Contains("funding") && Stocks[i].ToLower().Trim().StartsWith("910") && Accounts[i].ToLower().Trim().EndsWith("zaver"))
                {
                    isUnitedZaver = true;
                }
                ws.Cells[row, 1].Value = srNo;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 1], false, false);
                ws.Cells[row, 2].Value = Stocks[i];
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 2], false, false);
                ws.Cells[row, 3].Value = Accounts[i];
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 3], false, false);
                ws.Cells[row, 4].Value = Convert.ToDouble(Quantities[i]);
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 4], false, true);
                ws.Cells[row, 5].Value = Convert.ToDouble(Values[i]);
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 5], false, true);
                ws.Cells[row, 6].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 6], false, true);
                ws.Cells[row, 7].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 7], false, true);
                ws.Cells[row, 8].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 8], false, true);
                ws.Cells[row, 9].Value = "-";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 9], false, true);

                //I Will Call Python Api Here if i Don't get any values i will put the formula over here =K4*1000/N4
                double marketPrice = await Scrapping(Stocks[i].Split(' ')[0]);
                if (marketPrice == 0)
                {
                    ws.Cells[row, 10].Formula = $"K{row}*1000/N{row}";
                }
                else
                {
                    ws.Cells[row, 10].Value = marketPrice;

                }
                ////Changing the current month No.Of stock quantity if the key matches for the finalvalue of stocks;
                //string finalValueAfterSoldStockKey = $"Stock:{Stocks[i]}Account:{Accounts[i]}Amount:{Values[i]}";
                //if (FinalValuesAfterSoldStocks.ContainsKey(finalValueAfterSoldStockKey))
                //{
                //    Quantities[i] = FinalValuesAfterSoldStocks[finalValueAfterSoldStockKey].ToString();
                //}
                string finalValueAfterSoldStockKey = $"Stock:{Stocks[i]}Account:{Accounts[i]}";
                if (FinalValuesAfterSoldStocks.ContainsKey(finalValueAfterSoldStockKey))
                {
                    string tempQuantity = Quantities[i];
                    Values[i] = Convert.ToString(FinalValuesAfterSoldStocks[finalValueAfterSoldStockKey][0]);
                    Quantities[i] = Convert.ToString(FinalValuesAfterSoldStocks[finalValueAfterSoldStockKey][1]);
                    PurchaseValues[i] =Convert.ToString(Convert.ToDouble(Quantities[i]) * Convert.ToDouble(PurchaseValues[i]) / Convert.ToDouble(tempQuantity));

                }

                StockLineItemsFormatterCustomFormat(ws.Cells[row, 10], false, true);
                if (isUnited)
                {

                    ws.Cells[row, 11].Value = 0;
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
                    ws.Cells[row, 15].Value = 0;
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], false, true);
                }
                else if (isUSLineItems)
                {
                    ws.Cells[row, 11].Value = Convert.ToDouble(Values[i]);
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
                    ws.Cells[row, 15].Value = Convert.ToDouble(PurchaseValues[i]);
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], false, true);

                }
                else if (isUnitedZaver)
                {
                    ws.Cells[row, 11].Value = 0;
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], false, true);
                    ws.Cells[row, 15].Value = Convert.ToDouble(PurchaseValues[i]);
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], false, true);
                }
                else
                {
                    ws.Cells[row, 11].Formula = $"=J{row}*N{row}/1000";
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);


                    ws.Cells[row, 15].Value = Convert.ToDouble(PurchaseValues[i]);
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], false, true);

                }

                ws.Cells[row, 12].Formula = $"=K{row}-E{row}-F{row}";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 12], false, true);
                ws.Cells[row, 13].Formula = $"=IFERROR(L{row}/E{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 13], false, true);
                ws.Cells[row, 14].Value = Convert.ToDouble(Quantities[i]);
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 14], false, true);
                ws.Cells[row, 16].Formula = $"=K{row}-O{row}";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 16], false, true);
                ws.Cells[row, 17].Formula = $"=IFERROR(P{row}/O{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 17], false, true);
                row++;
                srNo++;

            }
            //Add Bought Stocks and unrealised Totals Cash and then total of all
            for (int i = 0; i < BoughtStocks.Count; i++)
            {
                bool isUSLineItems = false;
                if (Stocks[i].ToLower().Trim().Contains("u s") && Stocks[i].ToLower().Trim().StartsWith("9"))
                {
                    isUSLineItems = true;
                }
                ws.Cells[row, 1].Value = srNo;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 1], false, false);
                ws.Cells[row, 2].Value = BoughtStocks[i];
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 2], false, false);
                ws.Cells[row, 3].Value = BoughtAccounts[i];
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 3], false, false);
                ws.Cells[row, 4].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 4], false, true);
                ws.Cells[row, 5].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 5], false, true);
                ws.Cells[row, 6].Value = Convert.ToDouble(BoughtValues[i]);
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 6], false, true);
                ws.Cells[row, 7].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 7], false, true);
                ws.Cells[row, 8].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 8], false, true);
                ws.Cells[row, 9].Value = "-";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 9], false, true);
                if (isUSLineItems)
                {

                    ws.Cells[row, 10].Formula = $"=K{row}*1000/N{row}";
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 10], false, true);
                }
                else
                {
                    //I Will Call Python Api Here if i Don't get any values i will put the formula over here =K4*1000/N4
                    double marketPrice = await Scrapping(BoughtStocks[i].Split(' ')[0]);
                    if (marketPrice == 0)
                    {
                        ws.Cells[row, 10].Formula = $"=K{row}*1000/N{row}";
                    }
                    else
                    {

                        ws.Cells[row, 10].Value = marketPrice;
                    }
                    StockLineItemsFormatterCustomFormat(ws.Cells[row, 10], false, true);
                }
                ws.Cells[row, 11].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
                ws.Cells[row, 12].Formula = $"=K{row}-E{row}-F{row}";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 12], false, true);
                ws.Cells[row, 13].Formula = $"=IFERROR(L{row}/E{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 13], false, true);
                ws.Cells[row, 14].Value = Convert.ToDouble(BoughtQuantities[i]);
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 14], false, true);
                ws.Cells[row, 15].Value = Convert.ToDouble(BoughtValues[i]);
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], false, true);
                ws.Cells[row, 16].Formula = $"K{row} - O{row}";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 16], false, true);
                ws.Cells[row, 17].Formula = $"=IFERROR(P{row}/O{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 17], false, true);

                row++;
                srNo++;

            }

            //Total (Unrealised) Row
            ws.Cells[row, 1].Value = srNo;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 1], false, true);
            ws.Cells[row, 2].Value = "Total (Unrealised)";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 2], true, false);
            ws.Cells[row, 5].Formula = $"=SUM(E4:E{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 5], true, true);
            ws.Cells[row, 6].Formula = $"=SUM(F4:F{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 6], true, true);
            ws.Cells[row, 7].Formula = $"=SUM(G4:G{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 7], true, true);
            ws.Cells[row, 8].Formula = $"=SUM(H4:H{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 8], true, true);
            ws.Cells[row, 11].Formula = $"=SUM(K4:K{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
            ws.Cells[row, 12].Formula = $"=SUM(L4:L{row - 1})";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 12], true, true);
            ws.Cells[row, 13].Formula = $"=IFERROR(L{row}/E{row}*100,0)";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 13], true, true);
            ws.Cells[row, 15].Formula = $"=SUM(O4:O{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], true, true);
            ws.Cells[row, 16].Formula = $"=SUM(P4:P{row - 1})";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 16], true, true);
            ws.Cells[row, 17].Formula = $"=IFERROR(P{row - 1}/O{row - 1}*100,0)";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 17], true, true);


            //UnrealizedLoss Row and Columns are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("UnrealizedLoss", new List<int>());
            ConsolidateSheetFormulaValues["UnrealizedLoss"].Add(row);
            ConsolidateSheetFormulaValues["UnrealizedLoss"].Add(12);
            row++;
            srNo++;

            foreach (var kvp in SoldStocks)
            {

                string key = kvp.Key;
                double amount = kvp.Value;
                string stockPattern = @"Stock:(.*?)Account:";
                string accountPattern = @"Account:(.*?)OriginalPurchase:";
                string purchasePattern = @"OriginalPurchase:(.*?)StockPrevValues:";
                string prevMonthStockPattern = @"StockPrevValues:(.*?)Quantity:";
                string quantityPattern = @"Quantity:(\d+)";

                string stockName = GetStringFromRegexPattern(key, stockPattern);
                string accountName = GetStringFromRegexPattern(key, accountPattern);
                double quantity = Convert.ToDouble(GetStringFromRegexPattern(key, quantityPattern));
                double purchaseValue = Convert.ToDouble(GetStringFromRegexPattern(key, purchasePattern));
                double prevMonthStockValue = Convert.ToDouble(GetStringFromRegexPattern(key, prevMonthStockPattern));

                ws.Cells[row, 1].Value = srNo;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 1], false, false);
                ws.Cells[row, 2].Value = stockName;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 2], false, false);
                ws.Cells[row, 3].Value = accountName;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 3], false, false);

                ws.Cells[row, 4].Value = quantity;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 4], false, true);
                //Value as on previous Month
                ws.Cells[row, 5].Value = prevMonthStockValue;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 5], false, true);
                ws.Cells[row, 6].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 6], false, true);
                ws.Cells[row, 7].Value = -amount;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 7], false, true);
                ws.Cells[row, 8].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 8], false, true);
                ws.Cells[row, 9].Formula = $"=-G{row}-E{row}";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 9], false, true);
                ws.Cells[row, 10].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 10], false, true);
                ws.Cells[row, 11].Value = 0;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
                ws.Cells[row, 12].Value = 0;
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 12], false, true);
                ws.Cells[row, 13].Formula = $"=IFERROR(L{row}/E{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 13], false, true);
                ws.Cells[row, 14].Value = quantity;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 14], false, true);

                //Need to get the original value of purchase amount for this stock;
                ws.Cells[row, 15].Value = purchaseValue;
                StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], false, true);
                ws.Cells[row, 16].Formula = $"=-G{row}-O{row}";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 16], false, true);
                ws.Cells[row, 17].Formula = $"=IFERROR(P{row}/O{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 17], false, true);
                row++;
                srNo++;
            }
            //Total Unrealised+Realised Row
            ws.Cells[row, 1].Value = srNo;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 1], false, false);
            ws.Cells[row, 2].Value = "Total (Unrealised + Realised)";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 2], true, false);
            ws.Cells[row, 5].Formula = $"=SUM(E{row - 2}:E{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 5], true, true);
            ws.Cells[row, 6].Formula = $"=F{row - 2}";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 6], true, true);
            ws.Cells[row, 7].Formula = $"=SUM(G{row - 1}:G{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 7], true, true);
            ws.Cells[row, 8].Value = 0;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 8], true, true);
            ws.Cells[row, 9].Formula = $"=SUM(I{row - 1}:I{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 9], true, true);
            ws.Cells[row, 11].Formula = $"=SUM(K{row - 2}:K{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
            ws.Cells[row, 12].Formula = $"=K{row}-E{row}";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 12], true, true);
            ws.Cells[row, 13].Formula = $"=IFERROR(L{row}/E{row}*100,0)";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 13], true, true);
            ws.Cells[row, 15].Formula = $"=SUM(O{row - 2}:O{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 15], true, true);
            ws.Cells[row, 16].Formula = $"=SUM(P{row - 2}:P{row - 1})";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 16], true, true);
            ws.Cells[row, 17].Formula = $"=IFERROR(P{row}/O{row}*100,0)";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 17], true, true);

            //Empty row and columns also should have borders.
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 3], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 4], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 14], true, true);



            //OpeningPortfolioStock Row and Columns are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("OpeningPortfolioStock", new List<int>());
            ConsolidateSheetFormulaValues["OpeningPortfolioStock"].Add(row);
            ConsolidateSheetFormulaValues["OpeningPortfolioStock"].Add(5);

            //ClosingPortfolioStock Row and Columns are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("ClosingPortfolioStock", new List<int>());
            ConsolidateSheetFormulaValues["ClosingPortfolioStock"].Add(row);
            ConsolidateSheetFormulaValues["ClosingPortfolioStock"].Add(11);

            //RealizedGain Row and columns are added for getting the values in Consolidated sheet.
            ConsolidateSheetFormulaValues.Add("RealizedGain", new List<int>());
            ConsolidateSheetFormulaValues["RealizedGain"].Add(row);
            ConsolidateSheetFormulaValues["RealizedGain"].Add(9);

            //LTDPortfolioROI$ Row and columns are added for getting the values in Consolidated sheet.
            ConsolidateSheetFormulaValues.Add("LTDPortfolioROI$", new List<int>());
            ConsolidateSheetFormulaValues["LTDPortfolioROI$"].Add(row);
            ConsolidateSheetFormulaValues["LTDPortfolioROI$"].Add(16);

            //LTDPortfolioROI% Row and columns are added for getting the values in Consolidated sheet.
            ConsolidateSheetFormulaValues.Add("LTDPortfolioROI%", new List<int>());
            ConsolidateSheetFormulaValues["LTDPortfolioROI%"].Add(row);
            ConsolidateSheetFormulaValues["LTDPortfolioROI%"].Add(15);

            row++;
            srNo++;

            //Cash row 
            ws.Cells[row, 1].Value = srNo;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 1], false, false);
            ws.Cells[row, 2].Value = "Cash";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 2], true, false);

            //Cash value pull from the prev month table here 
            ws.Cells[row, 5].Value = cashValueCurrentMonth;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 5], true, true);

            ws.Cells[row, 6].Formula = $"=-F{row - 1}";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 6], true, true);
            ws.Cells[row, 7].Formula = $"=-G{row - 1}";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 7], true, true);
            //ACH Funds Value should be placed here
            ws.Cells[row, 8].Value = transferValue;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 8], true, true);

            //Profit and loss value like Divident+interest-Fee/1000 values placed here got all the data from the lpl file.
            ws.Cells[row, 9].Value = profitAndLossValue;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 9], true, true);

            ws.Cells[row, 10].Value = 0;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 10], true, true);
            ws.Cells[row, 11].Formula = $"=E{row}+F{row}+G{row}+H{row}+I{row}+J{row}";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
            ws.Cells[row, 12].Formula = $"=K{row}-E{row}";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 12], true, true);

            //Empty row and columns also should have borders.
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 3], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 4], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 13], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 14], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 15], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 16], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 17], true, true);

            //OpeningPortfolioCash Row and Columns are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("OpeningPortfolioCash", new List<int>());
            ConsolidateSheetFormulaValues["OpeningPortfolioCash"].Add(row);
            ConsolidateSheetFormulaValues["OpeningPortfolioCash"].Add(5);

            //ClosingPortfolio Row and Columns are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("ClosingPortfolioCash", new List<int>());
            ConsolidateSheetFormulaValues["ClosingPortfolioCash"].Add(row);
            ConsolidateSheetFormulaValues["ClosingPortfolioCash"].Add(11);

            //Sold Row and Columns  are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("SoldCash", new List<int>());
            ConsolidateSheetFormulaValues["SoldCash"].Add(row);
            ConsolidateSheetFormulaValues["SoldCash"].Add(7);

            //PurchaseCash row and Columns Are added for getitng the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("PurchaseCash", new List<int>());
            ConsolidateSheetFormulaValues["PurchaseCash"].Add(row);
            ConsolidateSheetFormulaValues["PurchaseCash"].Add(6);

            row++;
            srNo++;


            // Total unrealise+realised+cash row 
            ws.Cells[row, 1].Value = srNo;
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 1], false, false);
            ws.Cells[row, 2].Value = "Total (Unrealised + Realised + Cash)";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 2], true, false);
            ws.Cells[row, 5].Formula = $"=E{row - 1}+E{row - 2}";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 5], true, true);
            ws.Cells[row, 11].Formula = $"=SUM(K{row - 2}:K{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, 11], true, true);
            ws.Cells[row, 12].Formula = $"=K{row}-E{row}";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 12], true, true);

            //Empty row and columns also should have borders.
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 3], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 4], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 5], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 6], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 7], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 8], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 9], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 10], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 13], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 14], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 15], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 16], true, true);
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, 17], true, true);

            //OpeningPortfolio Row and Columns are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("OpeningPortfolio", new List<int>());
            ConsolidateSheetFormulaValues["OpeningPortfolio"].Add(row);
            ConsolidateSheetFormulaValues["OpeningPortfolio"].Add(5);

            //ClosingPortfolio Row and Columns are added for Getting the values in Consolidated Sheet.
            ConsolidateSheetFormulaValues.Add("ClosingPortfolio", new List<int>());
            ConsolidateSheetFormulaValues["ClosingPortfolio"].Add(row);
            ConsolidateSheetFormulaValues["ClosingPortfolio"].Add(11);

            //Calculate the main table
            ws.Cells[1, 1, row, ws.Dimension.End.Column].Calculate();

            bool isStarted = false;

            for (int i = 1; i <= row; i++)
            {
                var values = ws.Cells[i, 2].Value;
                if (values == null)
                {
                    continue;
                }
                if (values.ToString().Contains("Total (Unrealised)"))
                {
                    isStarted = true;
                    continue;
                }
                if (!isStarted)
                {
                    continue;
                }
                if (!values.ToString().Equals("Total (Unrealised + Realised)") && isStarted)
                {
                    double profitLossValue = Convert.ToDouble(ws.Cells[i, 9].Value);
                    if (profitLossValue < 0)
                    {
                        if (!ConsolidateSheetValues.ContainsKey("RealizedLoss"))
                        {
                            ConsolidateSheetValues.Add("RealizedLoss", profitAndLossValue);
                        }
                        else
                        {
                            ConsolidateSheetValues["RealizedLoss"] += profitAndLossValue;

                        }
                    }
                    double ltdValues = Convert.ToDouble(ws.Cells[i, 16].Value);
                    if (ltdValues > 0)
                    {
                        if (!ConsolidateSheetValues.ContainsKey("ProfitOnSale"))
                        {
                            ConsolidateSheetValues.Add("ProfitOnSale", ltdValues);
                        }
                        else
                        {

                            ConsolidateSheetValues["ProfitOnSale"] += ltdValues;
                        }
                    }
                    else
                    {
                        if (!ConsolidateSheetValues.ContainsKey("LossOnSale"))
                        {
                            ConsolidateSheetValues.Add("LossOnSale", ltdValues);
                        }
                        else
                        {

                            ConsolidateSheetValues["LossOnSale"] += ltdValues;
                        }
                    }
                }
            }
            return Tuple.Create(ConsolidateSheetFormulaValues, ConsolidateSheetValues);
        }

        public static string GetStringFromRegexPattern(string input, string pattern)
        {
            string result = string.Empty;
            Match match = Regex.Match(input, pattern);
            if (match.Success)
            {
                result = match.Groups[1].Value;
            }
            return result;
        }
        public static int GetCashRowOfPrevTable(OfficeOpenXml.ExcelWorksheet ws)
        {
            int rowNumber = 1;
            for (int row = 1; row <= ws.Dimension.End.Row; row++)
            {
                var values = ws.Cells[row, 2].Value;
                if (values == null)
                {
                    continue;
                }
                if (values.ToString().ToLower().Trim().Equals("cash"))
                {
                    rowNumber = row;
                    break;
                }
            }
            ws.Cells[1, 1, rowNumber, ws.Dimension.End.Column].Calculate();
            return rowNumber;

        }
        public static void EmptyRowStyleForSideTable(int row, int col, OfficeOpenXml.ExcelWorksheet ws)
        {
            for (int i = col; i <= col + 10; i++)
            {
                StockLineItemsFormatterCustomFormat(ws.Cells[row, i], true, true);
            }
        }

        public static void CreateSideTable(OfficeOpenXml.ExcelWorksheet ws, int col, int row, DateTime currentDate, Tuple<List<string>, List<string>> StockData, List<string> BoughtStocks, List<string> BoughtAccounts, Dictionary<string,List< DateTime>> BoughtStocksDate)
        {
            List<string> Stocks = StockData.Item1;
            List<string> Accounts = StockData.Item2;
            int prevMonthEstRow = Stocks.Count + BoughtStocks.Count + 2;
            for (int i = 0; i < Stocks.Count; i++)
            {
                if (Stocks[i].ToLower().Contains("united") && Accounts[i].ToLower().EndsWith("dfl"))
                {
                    EmptyRowStyleForSideTable(row, col, ws);
                    row++;
                    continue;
                }
                //Adding Formulas in the side Table
                ws.Cells[row, col].Formula = $"=B{row}";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col], false, false);
                ws.Cells[row, col + 1].Formula = $"=O{row}";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 1], false, true);
                ws.Cells[row, col + 2].Formula = $"=N{row}";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 2], false, true);
                ws.Cells[row, col + 3].Formula = $"=U{row}*1000/N{row}";
                StockLineItemsTwoDecimalPlacesFormat(ws.Cells[row, col + 3], false);
                ws.Cells[row, col + 4].Formula = $"=INDEX(X{prevMonthEstRow}:X{ws.Dimension.End.Row}, MATCH(T{row}, T{prevMonthEstRow}:T{ws.Dimension.End.Row}, 0))";
                StockLineItemsDateFormat(ws.Cells[row, col + 4], false);
                ws.Cells[row, col + 5].Formula = $"=J{row}";
                StockLineItemsTwoDecimalPlacesFormat(ws.Cells[row, col + 5], false);
                ws.Cells[row, col + 6].Formula = $"=(Y{row}-W{row})*N{row}/1000";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 6], false, true);
                ws.Cells[row, col + 7].Value = currentDate;
                StockLineItemsDateFormat(ws.Cells[row, col + 7], false);
                ws.Cells[row, col + 8].Formula = $"=DATEDIF(X{row}, AA{row}, \"D\")";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 8], false, true);
                ws.Cells[row, col + 9].Formula = $"=IFERROR(Z{row}/U{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 9], false, true);
                ws.Cells[row, col + 10].Formula = $"=IFERROR(AC{row}/AB{row}*256,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 10], false, true);
                row++;
            }
            Dictionary<string, int> RecordCompositeKeys = new Dictionary<string, int>();

            for (int i = 0; i < BoughtStocks.Count; i++)
            {
                ws.Cells[row, col].Formula = $"=B{row}";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col], false, false);
                ws.Cells[row, col + 1].Formula = $"=O{row}";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 1], false, true);
                ws.Cells[row, col + 2].Formula = $"=N{row}";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 2], false, true);
                ws.Cells[row, col + 3].Formula = $"=U{row}*1000/N{row}";
                StockLineItemsTwoDecimalPlacesFormat(ws.Cells[row, col + 3], false);
                //Bought Stocks Purchase date should be here with matching Composite Key.
                string compositeKey = $"Stock:{BoughtStocks[i]}Account:{BoughtAccounts[i]}";
                if (!RecordCompositeKeys.ContainsKey(compositeKey))
                {
                    RecordCompositeKeys.Add(compositeKey, 0);
                }
                else
                {
                    RecordCompositeKeys[compositeKey]++;
                }

                //handle the duplicate exception using below line
                //var matchingEntries = BoughtStocksDate.Where(kvp => kvp.Key == compositeKey).FirstOrDefault();
                //ws.Cells[row, col + 4].Value = matchingEntries;
                  ws.Cells[row, col + 4].Value = BoughtStocksDate[compositeKey][RecordCompositeKeys[compositeKey]];

                StockLineItemsDateFormat(ws.Cells[row, col + 4], false);
                ws.Cells[row, col + 5].Formula = $"=J{row}";
                StockLineItemsTwoDecimalPlacesFormat(ws.Cells[row, col + 5], false);
                ws.Cells[row, col + 6].Formula = $"=(Y{row}-W{row})*N{row}/1000";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 6], false, true);
                ws.Cells[row, col + 7].Value = currentDate;
                StockLineItemsDateFormat(ws.Cells[row, col + 7], false);
                ws.Cells[row, col + 8].Formula = $"=DATEDIF(X{row}, AA{row}, \"D\")";
                StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 8], false, true);
                ws.Cells[row, col + 9].Formula = $"=IFERROR(Z{row}/U{row}*100,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 9], false, true);
                ws.Cells[row, col + 10].Formula = $"=IFERROR(AC{row}/AB{row}*256,0)";
                StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 10], false, true);
                row++;
            }

            //Doing this to calculate the days holding column so that i will updated yoy formulas based on days holdings.
            ws.Cells[$"T2:AB{row}"].Calculate();
            for (int i = 3; i < row; i++)
            {
                var values = ws.Cells[i, 20].Value;
                if (values == null || values.ToString().ToLower().Equals("stock"))
                {
                    continue;
                }
                //var x = ws.Cells[i, 28].Value;
                double daysHolding = Convert.ToDouble(ws.Cells[i, 28].Value);
                if (daysHolding > 365)
                {
                    ws.Cells[i, col + 10].Formula = $"=AC{i}/AB{i}*365";
                }
                else
                {
                    ws.Cells[i, col + 10].Formula = $"=AC{i}/AB{i}*AB{i}";
                }
            }

            //Total Row For Side Table
            ws.Cells[row, col].Value = "Total";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col], true, false);
            ws.Cells[row, col + 1].Formula = $"=SUM(U4:U{row - 1})";
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 1], false, false);
            ws.Cells[row, col + 6].Formula = $"=SUM(Z4:Z{row - 1})";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 6], false, false);
            ws.Cells[row, col + 9].Formula = $"=Z{row}*100/U{row}";
            StockLineItemsFormatterCustomRedFormat(ws.Cells[row, col + 9], false, false);

            //Empty row and columns also should have borders.
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 2], false, false);
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 3], false, false);
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 4], false, false);
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 5], false, false);
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 7], false, false);
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 8], false, false);
            StockLineItemsFormatterCustomFormat(ws.Cells[row, col + 10], false, false);

        }

        public static void TitleFormatter(ExcelRange excelRange)
        {
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Size = 22;
        }
        public static void SecondLineTitleFormatter(ExcelRange excelRange)
        {
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Italic = true;
            excelRange.Style.Font.Size = 16;

        }

        public static void TableColumnHeadersFormatter(ExcelRange excelRange)
        {
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Size = 15;
            excelRange.Style.WrapText = true;
            excelRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            excelRange.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 0));
            excelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            excelRange.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            excelRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);


        }
        public static string GetExcelColumnName(int columnNumber)
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

        public static void StockLineItemsFormatterCustomFormat(ExcelRange excelRange, bool isBold, bool isAlignment)
        {
            if (isBold)
            {
                excelRange.Style.Font.Bold = true;
            }
            if (isAlignment)
            {
                excelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            }
            excelRange.Style.Font.Size = 16;
            excelRange.Style.Numberformat.Format = @"_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)";
            excelRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

        }

        public static void TwoDecimalPlacesCustomFormat(ExcelRange excelRange)
        {

        }
        public static void StockLineItemsFormatterCustomRedFormat(ExcelRange excelRange, bool isBold, bool isAlignment)
        {
            if (isBold)
            {
                excelRange.Style.Font.Bold = true;
            }
            if (isAlignment)
            {
                excelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
            }
            excelRange.Style.Font.Size = 16;
            excelRange.Style.Numberformat.Format = @"0;[Red](0)";
            excelRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

        }
        public static void StockLineItemsTwoDecimalPlacesFormat(ExcelRange excelRange, bool isBold)
        {
            if (isBold)
            {
                excelRange.Style.Font.Bold = true;
            }

            excelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            excelRange.Style.Numberformat.Format = "0.00";
            excelRange.Style.Font.Size = 16;
            excelRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
        }
        public static void StockLineItemsDateFormat(ExcelRange excelRange, bool isBold)
        {
            if (isBold)
            {
                excelRange.Style.Font.Bold = true;
            }

            excelRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

            excelRange.Style.Numberformat.Format = "mm/dd/yyyy";
            excelRange.Style.Font.Size = 16;
            excelRange.Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
        }
        public static void CreateConsolidatedSheet(OfficeOpenXml.ExcelWorksheet ws, DateTime currentDate, Tuple<Dictionary<string, List<int>>, Dictionary<string, double>> MainTableTuple, Dictionary<string, double> LPLValues, bool isNewYear)
        {
            Dictionary<string, List<int>> ConsolidateFormulaValues = MainTableTuple.Item1;
            Dictionary<string, double> ConsolidateValues = MainTableTuple.Item2;

            //If it is new year i want to copy all the style based on the 4th column due to i have added 3rd column without any formatting.
            if (isNewYear)
            {
                ws.InsertColumn(3, 2, 4);

            }
            else
            {
                ws.InsertColumn(3, 2, 3);
            }
            //Notes Column Value
            ws.Cells[2, 4].Value = "Notes";
            //Hide the Notes Column
            ws.Column(4).Hidden = true;
            //Year of the current month .
            ws.Cells[1, 3].Value = currentDate.Year;
            //Current month Month and Day
            ws.Cells[2, 3].Value = $"{currentDate.ToString("MMM")} {currentDate.ToString("dd")}";

            //There are total 50 row line items;
            //1. Opening PortFolio Value From Total(Unrealized+Realize+Cash) prev month Value Column  

            ws.Cells[3, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["OpeningPortfolio"][1], ConsolidateFormulaValues["OpeningPortfolio"][0]);
            //ws.Cells[3, 3].Formula = "";

            //2. Closing PortFolio Value From Total(Unrealized+Realize+Cash) current month Value Column Form CurrentMonth Value Column.

            ws.Cells[4, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["ClosingPortfolio"][1], ConsolidateFormulaValues["ClosingPortfolio"][0]);
            //ws.Cells[4,3].Formula="";

            //3. Opening PortFolio Cash Get the opening cash value from the cash row and previous month value Column.
            ws.Cells[5, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["OpeningPortfolioCash"][1], ConsolidateFormulaValues["OpeningPortfolioCash"][0]);
            //ws.Cells[5,3].Formula="";

            //4. Closing Portfolio Cash Get the closign cash value from the cash row and current month value column.
            //ws.Cells[6,3].Formula="";
            ws.Cells[6, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["ClosingPortfolioCash"][1], ConsolidateFormulaValues["ClosingPortfolioCash"][0]);

            // 5. Opening Portfolio Stock Value Get the value from the (Total Unrealized + realized) row Prev Month Value Column.
            ws.Cells[7, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["OpeningPortfolioStock"][1], ConsolidateFormulaValues["OpeningPortfolioStock"][0]);

            // 6. Closing Portfolio Stock Value Get the value from the (Total Unrealized + realized) row Current Month Value Column.
            ws.Cells[8, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["ClosingPortfolioStock"][1], ConsolidateFormulaValues["ClosingPortfolioStock"][0]);

            // 7. Net Change Portfolio
            ws.Cells[9, 3].Formula = "C4-C3";

            // 8. Net Change Stocks
            ws.Cells[10, 3].Formula = "C8-C7";

            // 9. Net Change Cash
            ws.Cells[11, 3].Formula = "C6-C5";

            // 10. Net Stock  Portfolio Performance =SUM(C30:C31)+SUM(C35:C36)
            ws.Cells[12, 3].Formula = "=SUM(C30:C31)+SUM(C35:C36)";
            // 11. Net Portfolio Performance =(C20+C21+C27+C28+C30+C31+C35+C36+C41)
            ws.Cells[13, 3].Formula = "=SUM(C20+C21+C27+C28+C30+C31+C35+C36+C41)";
            // 12 . Net Portfolio Performance (Realized + Cash) =(C20+C21+C28+C31+C36)
            ws.Cells[14, 3].Formula = "=SUM(C20+C21+C28+C31+C36)";

            // 13. Total Cash In =SUM(C16:C21)
            ws.Cells[15, 3].Formula = "=SUM(C16:C21)";

            // 14. LOC In Dont Know
            ws.Cells[16, 3].Value = 0;

            // 15. Transfer In Don't know 
            ws.Cells[17, 3].Value = 0;

            // 16. Sold Values are getting from the Cash row Sale column
            ws.Cells[18, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["SoldCash"][1], ConsolidateFormulaValues["SoldCash"][0]);

            // 17. Outside Cash In Don't Know
            ws.Cells[19, 3].Value = 0;

            // 18. Interest Income Filter LPL file Activity Column Interests Values /1000
            if (LPLValues.ContainsKey("Interest"))
            {
                ws.Cells[20, 3].Value = LPLValues["Interest"];
            }
            else
            {
                ws.Cells[20, 3].Value = 0;
            }

            // 19. Dividends  Filter LPL file Activity Column Dividemds Values /1000
            if (LPLValues.ContainsKey("Dividend"))
            {
                ws.Cells[21, 3].Value = LPLValues["Dividend"];
            }
            else
            {
                ws.Cells[21, 3].Value = 0;
            }

            // 20. Total Cash Out =SUM(C23:C28)
            ws.Cells[22, 3].Formula = "=SUM(C23:C28)";

            // 21. LOC Out Don't know.
            ws.Cells[23, 3].Value = 0;

            // 22. Transfer Out Ach funds Lpl file activity filter column to ach funds .
            if (LPLValues.ContainsKey("TransferOut"))
            {
                ws.Cells[24, 3].Value = LPLValues["TransferOut"];
            }
            else
            {
                ws.Cells[24, 3].Value = 0;
            }

            // 23. Purchase  Cash Row Buy Column
            ws.Cells[25, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["PurchaseCash"][1], ConsolidateFormulaValues["PurchaseCash"][0]);

            // 24. Outside Cash out Don't know.
            ws.Cells[26, 3].Value = 0;

            // 25. WireFees Don't know.
            if (LPLValues.ContainsKey("Tax"))
            {
                ws.Cells[27, 3].Value = LPLValues["Tax"];
            }
            else
            {
                ws.Cells[27, 3].Value = 0;
            }

            // 26. Fees and Expenses Filter LPL file Activity Column Fee values /1000.
            if (LPLValues.ContainsKey("FeesExpenses"))
            {
                ws.Cells[28, 3].Value = LPLValues["FeesExpenses"];
            }
            else
            {
                ws.Cells[28, 3].Value = 0;
            }

            // 27. Total Stock Increase =SUM(C30:C33)
            ws.Cells[29, 3].Formula = "=SUM(C30:C33)";

            // 28. Unrealized Gain Dont know.
            ws.Cells[30, 3].Value = 0;

            // 29. Realized Gain Profit/loss column Total(Unrealized+realized) row.
            ws.Cells[31, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["RealizedGain"][1], ConsolidateFormulaValues["RealizedGain"][0]);

            // 30. Purchase should be postive of Cash Row Buy Column
            ws.Cells[32, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["PurchaseCash"][1], ConsolidateFormulaValues["PurchaseCash"][0], true);

            // 31. Transferred in Stokcs Don't know.
            ws.Cells[33, 3].Value = 0;

            // 32. Total Stock Decrease =SUM(C35:C38)
            ws.Cells[34, 3].Formula = "=SUM(C35:C38)";

            // 33. Unrealized Loss Unrealized Row Profit/loss Column
            ws.Cells[35, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["UnrealizedLoss"][1], ConsolidateFormulaValues["UnrealizedLoss"][0]);

            // 34. Realized Loss Bought Stocks Negative values of profit/loss column
            if (ConsolidateValues.ContainsKey("RealizedLoss"))
            {
                ws.Cells[36, 3].Value = ConsolidateValues["RealizedLoss"];
            }
            else
            {
                ws.Cells[36, 3].Value = 0;
            }

            // 35. Sold Sale value of cash row Sale column.
            ws.Cells[37, 3].Formula = CreateIndividualSheetFormula(ConsolidateFormulaValues["SoldCash"][1], ConsolidateFormulaValues["SoldCash"][0]);

            // 36. Transferred Out Stock Don't know 
            ws.Cells[38, 3].Value = 0;

            // 37. Profit on Sale of Stocks Over Purchase (LTD $) LTD Change ($) sold stocks all positive values.
            if (ConsolidateValues.ContainsKey("ProfitOnSale"))
            {
                ws.Cells[39, 3].Value = ConsolidateValues["ProfitOnSale"];
            }
            else
            {
                ws.Cells[39, 3].Value = 0;
            }

            // 38. Loss on Sale of Stocks Over Purchase (LTD $) LTD Change($) sold stokcs values all negative values.
            if (ConsolidateValues.ContainsKey("LossOnSale"))
            {
                ws.Cells[40, 3].Value = ConsolidateValues["LossOnSale"];
            }
            else
            {
                ws.Cells[40, 3].Value = 0;
            }

            // 39. Cost of Change Don't know .
            ws.Cells[41, 3].Value = 0;

            // 40. Stock (Realized + Unrealized) ROI % =C12/C7*100
            ws.Cells[42, 3].Formula = "=C12/C7*100";

            // 41. Stock (Realized + Unrealized) ROI $ =SUM(C12)
            ws.Cells[43, 3].Formula = "=SUM(C12)";

            // 42. Portfolio (Realized) ROI % =((C31+C20+C21+C28+C27+C41+C36)/C3)*100
            ws.Cells[44, 3].Formula = "=((C31+C20+C21+C28+C27+C41+C36)/C3)*100";

            // 43. Portfolio (Realized) ROI $ =SUM(C14)
            ws.Cells[45, 3].Formula = "=SUM(C14)";

            // 44. Portfolio (Realized + Unrealized + Cash) ROI %  =(C20+C21+C28+C27+C41+C12)/(C3)*100
            ws.Cells[46, 3].Formula = "=(C20+C21+C28+C27+C41+C12)/(C3)*100";

            // 45. Portfolio (Realized + Unrealized + Cash) ROI $  =(C20+C21+C28+C27+C41+C12)
            ws.Cells[47, 3].Formula = "=(C20+C21+C28+C27+C41+C12)";

            string firstMonthColLetter = FindFirstMonthColumnOfCurrentYear(ws, currentDate);

            // 46. YTD Stock (Realized + Unrealized) ROI % =SUM(C42:$I$42)
            ws.Cells[48, 3].Formula = $"=SUM(C42:${firstMonthColLetter}$42)";

            // 47. YTD Portfolio (Realized + Unrealized + Cash) ROI $  =SUM(C$47:$I47)
            ws.Cells[49, 3].Formula = $"=SUM(C$47:${firstMonthColLetter}47)";

            // 48. YTD Portfolio (Realized + Unrealized + Cash) ROI %  =SUM(C$46:$I46)
            ws.Cells[50, 3].Formula = $"=SUM(C$46:${firstMonthColLetter}46)";

            // 49. LTD Portfolio (Realized + Unrealized + Cash) ROI $ Total Unrealize + Realized row LTD chaneg$ column + ='Individual Stocks'!P21+(SUM(C$20:$I20)+SUM(C$21:$I21)+SUM(C$28:$I28)+SUM(C$41:$I41))+(SUM(O$20:$AK20)+SUM(O$21:$AK21)+SUM(O$28:$AK28)+SUM(O$41:$AK41))

            //string overallFirstMonthColLetter = OverallFirstMonthColumn(ws);
            ws.Cells[51, 3].Formula = $"{CreateIndividualSheetFormula(ConsolidateFormulaValues["LTDPortfolioROI$"][1], ConsolidateFormulaValues["LTDPortfolioROI$"][0])}+(SUM(C$20:${firstMonthColLetter}20)+SUM(C$21:${firstMonthColLetter}21)+SUM(C$28:${firstMonthColLetter}28)+SUM(C$41:${firstMonthColLetter}41)){CreateLTDFormula(ws, currentDate)}";

            // 50. LTD Portfolio (Realized + Unrealized + Cash) ROI %  =C51/'Individual Stocks'!O21*100 Original purchase value column and totalunrealized + realized row .s
            string ltdPortfolioColLetter = GetExcelColumnName(ConsolidateFormulaValues["LTDPortfolioROI%"][1]);
            int ltdPortfolioRow = ConsolidateFormulaValues["LTDPortfolioROI%"][0];
            ws.Cells[52, 3].Formula = $"C51/'Individual Stocks'!{ltdPortfolioColLetter}{ltdPortfolioRow}*100";
            int firstMonthCol = FindFirstMonthColumnNumberOfCurrentYear(ws, currentDate);
            FormulaUpdateInYTDColumn(ws, firstMonthCol + 2, currentDate);
            List<string> MonthColumns = GetAllMonthsOfCurrentYear(ws, currentDate);
            UpdateAverageFormulaForConsolidatedSheet(firstMonthCol + 4, MonthColumns, ws);
        }

        public static void FormulaUpdateInYTDColumn(OfficeOpenXml.ExcelWorksheet ws, int ytdCol, DateTime currentDate)
        {
            List<string> MonthColumns = GetAllMonthsOfCurrentYear(ws, currentDate);

            ws.Cells[3, ytdCol].Formula = $"{MonthColumns[MonthColumns.Count - 1]}3";
            ws.Cells[4, ytdCol].Formula = $"{MonthColumns[0]}4";
            ws.Cells[5, ytdCol].Formula = $"{MonthColumns[MonthColumns.Count - 1]}5";
            ws.Cells[6, ytdCol].Formula = $"{MonthColumns[0]}6";
            ws.Cells[7, ytdCol].Formula = $"{MonthColumns[MonthColumns.Count - 1]}7";
            ws.Cells[8, ytdCol].Formula = $"{MonthColumns[0]}8";
            ws.Cells[9, ytdCol].Formula = $"{MonthColumns[0]}4-{MonthColumns[MonthColumns.Count - 1]}3";
            ws.Cells[10, ytdCol].Formula = $"{MonthColumns[0]}8-{MonthColumns[MonthColumns.Count - 1]}7";
            ws.Cells[11, ytdCol].Formula = $"{MonthColumns[0]}6-{MonthColumns[MonthColumns.Count - 1]}5";
            ws.Cells[12, ytdCol].Formula = $"=SUM({MonthColumns[0]}12:{MonthColumns[MonthColumns.Count - 1]}12)";
            ws.Cells[13, ytdCol].Formula = $"=SUM({MonthColumns[0]}13:{MonthColumns[MonthColumns.Count - 1]}13)";
            ws.Cells[14, ytdCol].Formula = $"=SUM({MonthColumns[0]}14:{MonthColumns[MonthColumns.Count - 1]}14)";
            ws.Cells[15, ytdCol].Formula = $"=SUM({MonthColumns[0]}15:{MonthColumns[MonthColumns.Count - 1]}15)";
            ws.Cells[16, ytdCol].Formula = $"=SUM({MonthColumns[0]}16:{MonthColumns[MonthColumns.Count - 1]}16)";
            ws.Cells[17, ytdCol].Formula = $"=SUM({MonthColumns[0]}17:{MonthColumns[MonthColumns.Count - 1]}17)";
            ws.Cells[18, ytdCol].Formula = $"=SUM({MonthColumns[0]}18:{MonthColumns[MonthColumns.Count - 1]}18)";
            ws.Cells[19, ytdCol].Formula = $"=SUM({MonthColumns[0]}19:{MonthColumns[MonthColumns.Count - 1]}19)";
            ws.Cells[20, ytdCol].Formula = $"=SUM({MonthColumns[0]}20:{MonthColumns[MonthColumns.Count - 1]}20)";
            ws.Cells[21, ytdCol].Formula = $"=SUM({MonthColumns[0]}21:{MonthColumns[MonthColumns.Count - 1]}21)";
            ws.Cells[22, ytdCol].Formula = $"=SUM({MonthColumns[0]}22:{MonthColumns[MonthColumns.Count - 1]}22)";
            ws.Cells[23, ytdCol].Formula = $"=SUM({MonthColumns[0]}23:{MonthColumns[MonthColumns.Count - 1]}23)";
            ws.Cells[24, ytdCol].Formula = $"=SUM({MonthColumns[0]}24:{MonthColumns[MonthColumns.Count - 1]}24)";
            ws.Cells[25, ytdCol].Formula = $"=SUM({MonthColumns[0]}25:{MonthColumns[MonthColumns.Count - 1]}25)";
            ws.Cells[26, ytdCol].Formula = $"=SUM({MonthColumns[0]}26:{MonthColumns[MonthColumns.Count - 1]}26)";
            ws.Cells[27, ytdCol].Formula = $"=SUM({MonthColumns[0]}27:{MonthColumns[MonthColumns.Count - 1]}27)";
            ws.Cells[28, ytdCol].Formula = $"=SUM({MonthColumns[0]}28:{MonthColumns[MonthColumns.Count - 1]}28)";
            ws.Cells[29, ytdCol].Formula = $"=SUM({MonthColumns[0]}29:{MonthColumns[MonthColumns.Count - 1]}29)";
            ws.Cells[30, ytdCol].Formula = $"=SUM({MonthColumns[0]}30:{MonthColumns[MonthColumns.Count - 1]}30)";
            ws.Cells[31, ytdCol].Formula = $"=SUM({MonthColumns[0]}31:{MonthColumns[MonthColumns.Count - 1]}31)";
            ws.Cells[32, ytdCol].Formula = $"=SUM({MonthColumns[0]}32:{MonthColumns[MonthColumns.Count - 1]}32)";
            ws.Cells[33, ytdCol].Formula = $"=SUM({MonthColumns[0]}33:{MonthColumns[MonthColumns.Count - 1]}33)";
            ws.Cells[34, ytdCol].Formula = $"=SUM({MonthColumns[0]}34:{MonthColumns[MonthColumns.Count - 1]}34)";
            ws.Cells[35, ytdCol].Formula = $"=SUM({MonthColumns[0]}35:{MonthColumns[MonthColumns.Count - 1]}35)";
            ws.Cells[36, ytdCol].Formula = $"=SUM({MonthColumns[0]}36:{MonthColumns[MonthColumns.Count - 1]}36)";
            ws.Cells[37, ytdCol].Formula = $"=SUM({MonthColumns[0]}37:{MonthColumns[MonthColumns.Count - 1]}37)";
            ws.Cells[38, ytdCol].Formula = $"=SUM({MonthColumns[0]}38:{MonthColumns[MonthColumns.Count - 1]}38)";
            ws.Cells[39, ytdCol].Formula = $"=SUM({MonthColumns[0]}39:{MonthColumns[MonthColumns.Count - 1]}39)";
            ws.Cells[40, ytdCol].Formula = $"=SUM({MonthColumns[0]}40:{MonthColumns[MonthColumns.Count - 1]}40)";
            ws.Cells[41, ytdCol].Formula = $"=SUM({MonthColumns[0]}41:{MonthColumns[MonthColumns.Count - 1]}41)";
            ws.Cells[42, ytdCol].Formula = $"={MonthColumns[0]}48";
            ws.Cells[43, ytdCol].Formula = $"=SUM({MonthColumns[0]}43:{MonthColumns[MonthColumns.Count - 1]}43)";
            ws.Cells[44, ytdCol].Formula = $"=SUM({MonthColumns[0]}44:{MonthColumns[MonthColumns.Count - 1]}44)";
            ws.Cells[45, ytdCol].Formula = $"=SUM({MonthColumns[0]}45:{MonthColumns[MonthColumns.Count - 1]}45)";
            ws.Cells[46, ytdCol].Formula = $"={MonthColumns[0]}50";
            ws.Cells[47, ytdCol].Formula = $"=SUM({MonthColumns[0]}47:{MonthColumns[MonthColumns.Count - 1]}47)";
            ws.Cells[48, ytdCol].Formula = $"=K12/K7*100";
            ws.Cells[49, ytdCol].Formula = $"{MonthColumns[0]}49";
            ws.Cells[50, ytdCol].Formula = $"{MonthColumns[0]}50";
            ws.Cells[51, ytdCol].Formula = $"{MonthColumns[0]}51";
            ws.Cells[52, ytdCol].Formula = $"{MonthColumns[0]}52";
        }

        public static List<string> GetAllMonthsOfCurrentYear(OfficeOpenXml.ExcelWorksheet ws, DateTime currentDate)
        {
            List<string> MonthColumns = new List<string>();
            string FirstMonth = "Jan";
            int lastCol = 0;

            while (currentDate.ToString("MMM") != FirstMonth)
            {
                for (int col = 1; col <= ws.Dimension.End.Column; col++)
                {
                    var yearRow = ws.Cells[1, col].Value;
                    var monthRow = ws.Cells[2, col].Value;
                    if (yearRow == null || monthRow == null)
                    {
                        continue;
                    }
                    if (yearRow.ToString().Contains(currentDate.Year.ToString()) && monthRow.ToString().Contains(currentDate.ToString("MMM")))
                    {
                        MonthColumns.Add(GetExcelColumnName(col));
                        lastCol = col;
                        break;
                    }
                }
                currentDate = currentDate.AddMonths(-1);
            }
            MonthColumns.Add(GetExcelColumnName(lastCol + 2));
            return MonthColumns;
        }

        public static void UpdateAverageFormulaForConsolidatedSheet(int avgCol, List<string> MonthLetters, OfficeOpenXml.ExcelWorksheet ws)
        {
            int endRow = 0;
            for (int row = 3; row <= ws.Dimension.End.Row; row++)
            {
                if(ws.Cells[row, 1].Value == null)
                {
                    endRow = row - 1; 
                    break;
                }
            }
            for (int row = 3; row <= endRow; row++)
            {
                string averageFormula = $"=AVERAGE({CreateAverageFormulaForConsolidatedSheet(MonthLetters, row)})";
                ws.Cells[row, avgCol].Formula = averageFormula;
            }
        }
        public static string CreateAverageFormulaForConsolidatedSheet(List<string> MonthLetters, int rowNumber)
        {
            string result = string.Empty;
            for (int i = 0; i < MonthLetters.Count; i++)
            {
                if (i + 1 == MonthLetters.Count)
                {
                    result += $"{MonthLetters[i]}{rowNumber}";
                }
                else
                {
                    result += $"{MonthLetters[i]}{rowNumber},";
                }
            }
            return result;
        }

        public static void GroupPreviousYearMonthsConsolidatedSheet(OfficeOpenXml.ExcelWorksheet ws, DateTime currentDate)
        {
            string LastMonth = "Jan";
            int firstCol = 0;
            int lastCol = 0;
            bool isFirstCol = false;
            //Insert a column without any style;
            ws.InsertColumn(3, 1);

            //Insert 3 columns here for YTD,Notes And Average as per the year changes.
            ws.InsertColumn(3, 3, 5);

            ws.Cells[1, 3].Value = currentDate.Year;
            ws.Cells[2, 3].Value = "YTD";

            ws.Cells[2, 4].Value = "Notes";

            ws.Cells[2, 5].Value = "Average";

            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var values = ws.Cells[2, col].Value;
                var yearValues = ws.Cells[1, col].Value;
                if (values == null || yearValues == null)
                {
                    continue;
                }
                if (yearValues.ToString().Contains(currentDate.AddYears(-1).Year.ToString()) && !isFirstCol)
                {
                    isFirstCol = true;
                    firstCol = col;
                }

                if (values.ToString().Contains(LastMonth) && yearValues.ToString().Contains(currentDate.AddYears(-1).Year.ToString()))
                {
                    lastCol = col + 4;
                    break;
                }



            }
            for (int i = firstCol; i <= lastCol; i++)
            {
                ws.Column(i).OutlineLevel = 1;
                ws.Column(i).Collapsed = true;
                ws.Column(i).Hidden = true;
            }

        }

        public static string CreateIndividualSheetFormula(int col, int row, bool isNegative = false)
        {
            string colLetter = GetExcelColumnName(col);
            string result = isNegative ? $"=-'Individual Stocks'!{colLetter}{row}" : $"='Individual Stocks'!{colLetter}{row}";
            return result;

        }

        public static string FindFirstMonthColumnOfCurrentYear(OfficeOpenXml.ExcelWorksheet ws, DateTime currentDate)
        {
            string result = string.Empty;
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var monthValues = ws.Cells[2, col].Value;
                var yearValues = ws.Cells[1, col].Value;
                if (monthValues == null || yearValues == null)
                {
                    continue;
                }
                if (monthValues.ToString().Contains("Jan") && yearValues.ToString().Contains(currentDate.Year.ToString()))
                {
                    result = GetExcelColumnName(col);
                    break;
                }

            }
            return result;
        }

        public static int FindFirstMonthColumnNumberOfCurrentYear(OfficeOpenXml.ExcelWorksheet ws, DateTime currentDate)
        {
            int result = 0;
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var monthValues = ws.Cells[2, col].Value;
                var yearValues = ws.Cells[1, col].Value;
                if (monthValues == null || yearValues == null)
                {
                    continue;
                }
                if (monthValues.ToString().Contains("Jan") && yearValues.ToString().Contains(currentDate.Year.ToString()))
                {
                    result = col;
                    break;
                }

            }
            return result;
        }
        public static string OverallFirstMonthYear(OfficeOpenXml.ExcelWorksheet ws)
        {
            string year = string.Empty;
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var monthValues = ws.Cells[2, col].Value;
                if (monthValues == null)
                {
                    continue;
                }
                if (monthValues.ToString().Contains("Jan"))
                {

                    year = ws.Cells[1, col].Value.ToString();
                }
            }
            return year;
        }

        public static string CreateLTDFormula(OfficeOpenXml.ExcelWorksheet ws, DateTime currentDate)
        {
            string result = string.Empty;
            string overallLastYear = OverallFirstMonthYear(ws);
            while (currentDate.Year.ToString() != overallLastYear)
            {
                currentDate = currentDate.AddYears(-1);
                var columnTuple = GetFirstAndLastMonthOfYear(ws, currentDate);
                string firstMonth = columnTuple.Item1;
                string lastMonth = columnTuple.Item2;
                result += $"+(SUM({lastMonth}$20:${firstMonth}20)+SUM({lastMonth}$21:${firstMonth}21)+SUM({lastMonth}$28:${firstMonth}28)+SUM({lastMonth}$41:${firstMonth}41))";
            }
            return result;
        }

        public static Tuple<string, string> GetFirstAndLastMonthOfYear(OfficeOpenXml.ExcelWorksheet ws, DateTime currentDate)
        {
            string firstMonth = string.Empty;
            string lastMonth = string.Empty;
            for (int col = 1; col <= ws.Dimension.End.Column; col++)
            {
                var monthValue = ws.Cells[2, col].Value;
                var yearValue = ws.Cells[1, col].Value;
                if (monthValue == null || yearValue == null)
                {
                    continue;
                }
                if (monthValue.ToString().Contains("Jan") && yearValue.ToString().Contains(currentDate.Year.ToString()))
                {
                    firstMonth = GetExcelColumnName(col);
                }
                else if (monthValue.ToString().Contains("Dec") && yearValue.ToString().Contains(currentDate.Year.ToString()))
                {
                    lastMonth = GetExcelColumnName(col);
                }
            }
            return Tuple.Create(firstMonth, lastMonth);
        }
    }
}

