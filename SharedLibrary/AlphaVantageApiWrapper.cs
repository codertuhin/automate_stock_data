using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace AlphaVantageApiWrapper
{
    public static partial class AlphaVantageApiWrapper
    {
        public static async Task<AlphaVantageRootObject> GetTechnical(List<ApiParam> parameters, string apiKey)
        {
            var stringRequest = parameters.Aggregate(@"https://www.alphavantage.co/query?", (current, param) => current + param.ToApiString());
            stringRequest += "&apikey=" + apiKey;

            var apiData = await CallAlphaVantageApi(stringRequest);

            var technicalsObject = new AlphaVantageRootObject
            {
                MetaData = new MetaData
                {
                    Function = parameters.FirstOrDefault(x => x.ParamName.Equals("function"))?.ParamValue ?? "NA?",
                    Interval = parameters.FirstOrDefault(x => x.ParamName.Equals("interval"))?.ParamValue ?? "NA?",
                    SeriesType = parameters.FirstOrDefault(x => x.ParamName.Equals("series_type"))?.ParamValue ?? "NA?",
                    Symbol = parameters.FirstOrDefault(x => x.ParamName.Equals("symbol"))?.ParamValue ?? "NA?"
                },

                TechnicalsByDate = apiData.Last.Values().OfType<JProperty>().Select(x => new TechnicalDataDate
                {
                    Date = Convert.ToDateTime(x.Name),
                    Data = x.Value.OfType<JProperty>().Select(r => new TechnicalDataObject
                    {
                        TechnicalKey = r.Name,
                        TechnicalValue = Convert.ToDouble(r.Value.ToString())
                    }).ToList()
                })
                    .ToList()
            };

            return technicalsObject;
        }

        public static string ToDescription(this Enum enumeration)
        {
            var type = enumeration.GetType();
            var memInfo = type.GetMember(enumeration.ToString());
            if (memInfo.Length <= 0) return enumeration.ToString();
            var attrs = memInfo[0].GetCustomAttributes(typeof(EnumDescription), false);
            return attrs.Length > 0 ? ((EnumDescription)attrs[0]).Text : enumeration.ToString();
        }

        public static async Task<JObject> CallAlphaVantageApi(string stringRequest)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    var res = await client.GetStringAsync(stringRequest);
                    return JsonConvert.DeserializeObject<JObject>(res);
                }
            }
            catch (Exception e)
            {
                //fatal error
                return null;
            }
        }
    }
}