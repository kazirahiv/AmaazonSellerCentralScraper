using System;
using System.Collections.Generic;

namespace AmazonScraper.Models
{
    public class ScrapeData
    {
        public ScrapeData()
        {
            Reports = new List<Report>();
        }
        public List<Report> Reports { get; set; }
        public DateTime LastScraped { get; set; }
        public DateTime LastScrapedDatePickerTime { get; set; }
        public bool AllScrapedTillDate { get; set; }
    }
}
