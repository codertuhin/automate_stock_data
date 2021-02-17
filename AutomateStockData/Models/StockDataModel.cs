﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace AutomateStockData.Models
{
    public class StockDataModel
    {
        public DateTime Date { get; set; }
        public double Open { get; set; }
        public double High { get; set; }
        public double Low { get; set; }
        public double Close { get; set; }

    }
}