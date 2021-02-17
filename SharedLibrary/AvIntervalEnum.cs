namespace AlphaVantageApiWrapper
{
    public static partial class AlphaVantageApiWrapper
    {
        public enum AvIntervalEnum
        {
            [EnumDescription("1min")] OneMinute,
            [EnumDescription("5min")] FiveMinutes,
            [EnumDescription("15min")] FifteenMinutes,
            [EnumDescription("30min")] ThirtyMinutes,
            [EnumDescription("60min")] SixtyMinutes,
            [EnumDescription("daily")] Daily,
            [EnumDescription("weekly")] Weekly,
            [EnumDescription("monthly")] Monthly
        }
    }
}