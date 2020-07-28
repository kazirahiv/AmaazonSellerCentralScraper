using System;

namespace AmazonScraper.Models
{
    public class CollectionHistory
    {
        public DateTime LastScrapedTime { get; set; }
        public DateTime LastScrapedDatePickerTime { get; set; }
        public CollectionHistory(DateTime LastScrapedTime, DateTime LastScrapedDatePickerTime)
        {
            this.LastScrapedDatePickerTime = LastScrapedDatePickerTime;
            this.LastScrapedTime = LastScrapedTime;
        }

    }
}
