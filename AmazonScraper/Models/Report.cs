using System;

namespace AmazonScraper.Models
{
    public class Report
    {
        public string ParentASIN { get; set; }
        public string ChildASIN { get; set; }
        public string Sessions { get; set; }
        public string UnitsOrdered { get; set; }
        public string ProductSales { get; set; }
        public string TotalOrderItems { get; set; }
        public DateTime Date { get; set; }

        public override string ToString()
        {
            return ParentASIN + " " + ChildASIN + " " + Sessions + " " + UnitsOrdered + " " + ProductSales + " " + TotalOrderItems + "\n";
        }
    }
}
