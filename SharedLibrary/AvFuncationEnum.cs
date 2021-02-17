namespace AlphaVantageApiWrapper
{
    public static partial class AlphaVantageApiWrapper
    {
        public enum AvFuncationEnum
        {
            [EnumDescription("SMA")] Sma,
            [EnumDescription("EMA")] Ema,
            [EnumDescription("MACD")] Macd,
            [EnumDescription("STOCH")] Stoch,
            [EnumDescription("RSI")] Rsi,
        }
    }
}