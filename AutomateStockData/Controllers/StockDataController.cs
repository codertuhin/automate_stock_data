using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using YahooFinanceAPI;
using YahooFinanceAPI.Models;
using static AlphaVantageApiWrapper.AlphaVantageApiWrapper;

namespace AutomateStockData.Controllers
{
    public class StockDataController : ApiController
    {

        string csvdata = "";
        public async Task<IHttpActionResult> GetStockData()
        {
            //var data = GetHistoricalPrice("msft");
            //return Ok(data.Result);

            var API_KEY = "RJQACXXR1DL2WYJT";

            var StockTickers = new List<string> { "AAPL" };

            //foreach (var ticker in StockTickers)
            //{
            var parameters = new List<ApiParam>
                {
                    new ApiParam("function", AvFuncationEnum.Rsi.ToDescription()),
                    new ApiParam("symbol", "MSFT"),
                    new ApiParam("interval", AvIntervalEnum.OneMinute.ToDescription()),
                    new ApiParam("time_period", "5"),
                    new ApiParam("series_type", AvSeriesType.Open.ToDescription()),
                };

            //Start Collecting SMA values

            var SMA_5 = await GetTechnical(parameters, API_KEY);
            //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "20";
            //var SMA_20 = GetTechnical(parameters, API_KEY);
            //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "50";
            //var SMA_50 = GetTechnical(parameters, API_KEY);
            //parameters.FirstOrDefault(x => x.ParamName == "time_period").ParamValue = "200";
            //var SMA_200 = GetTechnical(parameters, API_KEY);





            using (var package = new ExcelPackage())
            {
                // Add a new worksheet to the empty workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("MSFT");
                //Add the headers
                worksheet.Cells[1, 1].Value = "Date";
                worksheet.Cells[1, 2].Value = "Technical Value";


                //Add some items...

                //foreach (var item in hps)
                //{
                //    worksheet.Cells[x, 1].Value = item.Date.ToString("MM/dd/yy");
                //    worksheet.Cells[x, 2].Value = item.Open;
                //    worksheet.Cells[x, 3].Value = item.High;
                //    worksheet.Cells[x, 4].Value = item.Low;
                //    worksheet.Cells[x, 5].Value = item.Close;
                //    worksheet.Cells[x, 6].Value = item.AdjClose;
                //    worksheet.Cells[x, 7].Value = item.Volume;


                //    x++;
                //}
                int x = 1;
                foreach (var item in SMA_5.TechnicalsByDate)
                {
                    foreach (var data in item.Data)
                    {
                        worksheet.Cells[x, 1].Value = item.Date.ToString("MM/dd/yy");
                        worksheet.Cells[x, 2].Value = data.TechnicalValue;
                        x++;
                    }
                }


                //worksheet.Calculate();

                worksheet.Cells.AutoFitColumns(0);  //Autofit columns for all cells




                var f = new FileInfo("test.xlsx");
                // save our new workbook in the output directory and we are done!
                package.SaveAs(f);

                return Ok(SMA_5);

                //}

                //return Ok();
            }
        }


        private async Task<List<HistoryPrice>> GetHistoricalPrice(string symbol)
        {

            //first get a valid token from Yahoo Finance
            while (string.IsNullOrEmpty(Token.Cookie) || string.IsNullOrEmpty(Token.Crumb))
            {
                await Token.RefreshAsync().ConfigureAwait(false);
            }

            List<HistoryPrice> hps = await Historical.GetPriceAsync(symbol, DateTime.Now.AddMonths(-1), DateTime.Now).ConfigureAwait(false);


            return hps;
            //do something

        }

        private async Task GetRawHistoricalPrice(string symbol)
        {

            //first get a valid token from Yahoo Finance
            while (string.IsNullOrEmpty(Token.Cookie) || string.IsNullOrEmpty(Token.Crumb))
            {
                await Token.RefreshAsync().ConfigureAwait(false);
            }

            csvdata = await Historical.GetRawAsync(symbol, DateTime.Now.AddMonths(-1), DateTime.Now).ConfigureAwait(false);

            //process further

        }
    }
}
