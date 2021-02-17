using System.Collections.Generic;

namespace AlphaVantageApiWrapper
{
    public static partial class AlphaVantageApiWrapper
    {
        public class AlphaVantageRootObject
        {
            public MetaData MetaData;
            public List<TechnicalDataDate> TechnicalsByDate;
        }
    }
}