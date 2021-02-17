using System;

namespace AlphaVantageApiWrapper
{
    public static partial class AlphaVantageApiWrapper
    {
        public class EnumDescription : Attribute
        {
            public string Text { get; }

            public EnumDescription(string text)
            {
                Text = text;
            }
        }
    }
}