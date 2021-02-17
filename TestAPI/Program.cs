using AlphaVantageApiWrapper;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using YahooFinanceAPI;
using YahooFinanceAPI.Models;
using static AlphaVantageApiWrapper.AlphaVantageApiWrapper;

namespace TestAPI
{
    class Program
    {
        static void Main(string[] args)
        {
            //GetHistoricalPrice().Wait();

            //GetRawHistoricalPrice("msft").Wait();

            //TestAsync().Wait();

            //GetStockData().Wait();

            List<string> lstSymbols = new List<string>() { "ABT", "SPY", "MSFT", "AAPL", "AXP", "AMT", "APTV", "VZ", "INTC", "IBM", "FB" };

            var data = Data(lstSymbols).Result;

            PrepareExcelSheet(data, lstSymbols, "1 Min - Open", "1-Min-Open.xlsx");

        }

        public static async Task<List<HistoryPrice>> Data(List<string> lstSymbols)
        {

            List<HistoryPrice> list = new List<HistoryPrice>();


            foreach (var symbol in lstSymbols)
            {
                string csvData = null;
                string uri = "https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={0}&interval=1min&apikey=RJQACXXR1DL2WYJT&datatype=csv";

                uri = string.Format(uri, symbol);

                using (var wc = new WebClient())
                {
                    wc.Headers.Add(HttpRequestHeader.Cookie, Token.Cookie);
                    csvData = await wc.DownloadStringTaskAsync(uri).ConfigureAwait(false);
                }


                var rows = csvData.Split(Convert.ToChar(10));

                //row(0) was ignored because is column names
                //data is read from oldest to latest
                for (var i = 1; i <= rows.Length - 1; i++)
                {
                    var row = rows[i];
                    if (string.IsNullOrEmpty(row)) continue;

                    var cols = row.Split(',');
                    if (cols[1] == "null") continue;

                    var itm = new HistoryPrice
                    {
                        Date = DateTime.Parse(cols[0]),
                        Open = Convert.ToDouble(cols[1]),
                        High = Convert.ToDouble(cols[2]),
                        Low = Convert.ToDouble(cols[3]),
                        Close = Convert.ToDouble(cols[4]),
                        Volume = Convert.ToDouble(cols[5]),
                        Symbol = symbol
                    };

                    //fixed issue in some currencies quote (e.g: SGDAUD=X)


                    list.Add(itm);

                }
            }
       
            

            return list;
        }

        //public async Task TSI()
        //{
        //    HttpClient client = new HttpClient();
        //    //Uri uri = new Uri("http://date.jsontest.com/");
        //    Uri uri = new Uri("https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=MSFT&interval=5min&apikey=7NIMRBR8G8UB7P8C");
        //    client.DefaultRequestHeaders.Accept.Clear();
        //    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        //    HttpResponseMessage response = await client.GetAsync(uri);
        //    if (response.IsSuccessStatusCode)
        //    {
        //        dynamic result = await response.Content.ReadAsStringAsync();

        //        IEnumerable<dynamic> dObj = JsonConvert.DeserializeObject<dynamic>(result.ToString());




        //        var jObj = JObject.Parse(json);
        //        var metadata = jObj["Meta Data"].ToObject<Dictionary<string, string>>();
        //        var timeseries = jObj["Time Series (1min)"].ToObject<Dictionary<string, Dictionary<string, string>>>();

        //    }
        //}


        public static async Task GetStockData()
        {
            //var data = GetHistoricalPrice("msft");
            //return Ok(data.Result);

            var API_KEY = "RJQACXXR1DL2WYJT";

            var StockTickers = new List<string> { "AAPL", "MSFT" };

            //foreach (var ticker in StockTickers)
            //{
            var parameters = new List<ApiParam>
                {
                    new ApiParam("function", AvFuncationEnum.Rsi.ToDescription()),
                    new ApiParam("symbol", "MSFT"),
                    new ApiParam("interval", AvIntervalEnum.OneMinute.ToDescription()),
                    new ApiParam("time_period", "1"),
                    new ApiParam("series_type", AvSeriesType.Open.ToDescription()),
                };

            //Start Collecting SMA values

            var SMA_12 = await GetTechnical(parameters, API_KEY);
            parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "14";
            var SMA_14 = await GetTechnical(parameters, API_KEY);
            //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "50";
            //var SMA_50 = GetTechnical(parameters, API_KEY);
            //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "200";
            //var SMA_200 = GetTechnical(parameters, API_KEY);





            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("MSFT RSI 12 Period");
                //Add the headers
                worksheet.Cells[1, 1].Value = "Date";
                worksheet.Cells[1, 2].Value = "Technical Value";


                int x = 2;
                foreach (var item in SMA_12.TechnicalsByDate)
                {
                    foreach (var data in item.Data)
                    {
                        worksheet.Cells[x, 1].Value = item.Date.ToString("MM/dd/yy HH:mm");
                        worksheet.Cells[x, 2].Value = data.TechnicalValue;
                        x++;
                    }
                }


                worksheet.Cells.AutoFitColumns(0);


                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("MSFT RSI 14 Period");
                //Add the headers
                worksheet2.Cells[1, 1].Value = "Date";
                worksheet2.Cells[1, 2].Value = "Technical Value";


                x = 2;
                foreach (var item in SMA_14.TechnicalsByDate)
                {
                    foreach (var data in item.Data)
                    {
                        worksheet2.Cells[x, 1].Value = item.Date.ToString("MM/dd/yy HH:mm");
                        worksheet2.Cells[x, 2].Value = data.TechnicalValue;
                        x++;
                    }
                }


                worksheet2.Cells.AutoFitColumns(0);


                var f = new FileInfo("test.xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(f);


                //}

                //return Ok();
            }
        }

        //RJQACXXR1DL2WYJT

        static async Task GetHistoricalPrice()
        {

            //first get a valid token from Yahoo Finance
            while (string.IsNullOrEmpty(Token.Cookie) || string.IsNullOrEmpty(Token.Crumb))
            {
                await Token.RefreshAsync().ConfigureAwait(false);
            }

            List<string> lstSymbols = new List<string>() {"ABT", "SPY", "MSFT", "AAPL", "AXP", "AMT", "APTV", "VZ", "INTC", "IBM", "FB" };

            List<HistoryPrice> list = new List<HistoryPrice>();

            foreach (var symbol in lstSymbols)
            {

                Console.WriteLine("Collecting Data for " + symbol + "....");
                List<HistoryPrice> hps = await Historical.GetPriceAsync(symbol, DateTime.Now.AddMonths(-1), DateTime.Now).ConfigureAwait(false);

                list.AddRange(hps);
            }




            //PrepareExcelSheet(list, lstSymbols);





            list.ForEach(Console.WriteLine);

            /*
            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(symbol.ToUpper());
                //Add the headers
                worksheet.Cells[1, 1].Value = "Date";
                worksheet.Cells[1, 2].Value = "Open";
                worksheet.Cells[1, 3].Value = "High";
                worksheet.Cells[1, 4].Value = "Low";
                worksheet.Cells[1, 5].Value = "Close";
                worksheet.Cells[1, 6].Value = "Adj. Close";
                worksheet.Cells[1, 7].Value = "Volume";

                //Add some items...
                int x = 2;
                foreach (var item in hps)
                {
                    worksheet.Cells[x, 1].Value = item.Date.ToString("MM/dd/yy");
                    worksheet.Cells[x, 2].Value = item.Open;
                    worksheet.Cells[x, 3].Value = item.High;
                    worksheet.Cells[x, 4].Value = item.Low;
                    worksheet.Cells[x, 5].Value = item.Close;
                    worksheet.Cells[x, 6].Value = item.AdjClose;
                    worksheet.Cells[x, 7].Value = item.Volume;


                    x++;
                }


                //worksheet.Calculate();

                worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells




                var f = new FileInfo("test.xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(f);

            }
    
            */


            //do something

            //Process.Start("test.xlsx");

            Console.WriteLine("Done...");
            Console.ReadLine();
        }

        static void PrepareExcelSheet(List<HistoryPrice> list, List<string> lstSymbols, string worksheetName, string fileName)
        {
            var data = from d in list
                       select new { d.Date, d.Open, d.Symbol };


            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);
                worksheet.Cells[1, 1].Value = "Date";

                int sym = 2;
                foreach (var symbol in lstSymbols)
                {

                    Console.WriteLine("Populating Data to Excel: {0}", symbol);
                    worksheet.Cells[1, sym].Value = symbol;

                    var symbolDataCollection = data.Where(v => v.Symbol == symbol);

                    int symbolDataCount = 2;
                    foreach (var symbolData in symbolDataCollection)
                    {
                        worksheet.Cells[symbolDataCount, 1].Value = symbolData.Date.ToString("dd/MMMM/yy HH:mm");
                        worksheet.Cells[symbolDataCount, sym].Value = symbolData.Open;

                        symbolDataCount++;
                    }

                    sym++;
                }


                var f = new FileInfo(fileName);
                package.SaveAs(f);

            }
        }

        static async Task GetRawHistoricalPrice(string symbol)
        {

            //first get a valid token from Yahoo Finance
            while (string.IsNullOrEmpty(Token.Cookie) || string.IsNullOrEmpty(Token.Crumb))
            {
                await Token.RefreshAsync().ConfigureAwait(false);
            }

            var csvdata = await Historical.GetRawAsync(symbol, DateTime.Now.AddYears(-1), DateTime.Now).ConfigureAwait(false);


            Console.WriteLine(csvdata);
            //process further

            Console.ReadLine();

        }

        public static async Task TestAsync()
        {
            var API_KEY = "RJQACXXR1DL2WYJT";

            var StockTickers = new List<string> { "AAPL" };

            foreach (var ticker in StockTickers)
            {
                var parameters = new List<ApiParam>
                {
                    new ApiParam("function", AvFuncationEnum.Sma.ToDescription()),
                    new ApiParam("symbol", ticker),
                    new ApiParam("interval", AvIntervalEnum.OneMinute.ToDescription()),
                    new ApiParam("time_period", "5"),
                    new ApiParam("series_type", AvSeriesType.Open.ToDescription()),
                };

                //Start Collecting SMA values

                var SMA_5 = await GetTechnical(parameters, API_KEY);
                parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "20";
                var SMA_20 = await GetTechnical(parameters, API_KEY);
                parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "50";
                var SMA_50 = await GetTechnical(parameters, API_KEY);
                parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "200";
                var SMA_200 = await GetTechnical(parameters, API_KEY);

                ////Change function to EMA
                //parameters.FirstOrDefault(x => x.ParamName == "function").ParamValue = AvFuncationEnum.Sma.ToDescription();

                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "5";
                //var EMA_5 = await GetTechnical(parameters, API_KEY);
                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "20";
                //var EMA_20 = await GetTechnical(parameters, API_KEY);
                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "50";
                //var EMA_50 = await GetTechnical(parameters, API_KEY);
                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "200";
                //var EMA_200 = await GetTechnical(parameters, API_KEY);

                //Change function to RSI
                //parameters.FirstOrDefault(x => x.ParamName == "function").ParamValue = AvFuncationEnum.Rsi.ToDescription();

                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "7";
                //var RSI_7 = await GetTechnical(parameters, API_KEY);
                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "14";
                //var RSI_14 = await GetTechnical(parameters, API_KEY);
                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "24";
                //var RSI_24 = await GetTechnical(parameters, API_KEY);
                //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "50";
                //var RSI_50 = await GetTechnical(parameters, API_KEY);

                //Change function to MACD
                //parameters.FirstOrDefault(x => x.ParamName == "function").ParamValue = AvFuncationEnum.Macd.ToDescription();
                ////Remove time period to use default values (slow 12, fast 26)
                //var itemToRemove = parameters.FirstOrDefault(x => x.ParamName == "time_period");
                //parameters.Remove(itemToRemove);
                //var MACD = await GetTechnical(parameters, API_KEY);

                ////Change function to STOCK
                //parameters.FirstOrDefault(x => x.ParamName == "function").ParamValue = AvFuncationEnum.Stoch.ToDescription();
                //var STOCH = await GetTechnical(parameters, API_KEY);
            }
        }
    }

    public class MyClass
    {

    }



}
