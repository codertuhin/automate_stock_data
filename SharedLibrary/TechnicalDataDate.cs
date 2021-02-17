using System;
using System.Collections.Generic;

namespace AlphaVantageApiWrapper
{
    public static partial class AlphaVantageApiWrapper
    {
        public class TechnicalDataDate
        {
            public DateTime Date;
            public List<TechnicalDataObject> Data;
        }
    }
}