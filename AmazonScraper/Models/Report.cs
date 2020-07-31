using System;

namespace AmazonScraper.Models
{
    public class Report
    {
        public string ParentASIN { get; set; }
        public string ChildASIN { get; set; }
        public int Sessions { get; set; }
        public int UnitsOrdered { get; set; }
        public decimal ProductSales { get; set; }
        public int TotalOrderItems { get; set; }
        public DateTime Date { get; set; }

        public override string ToString()
        {
            return ParentASIN + " " + ChildASIN + " " + Sessions + " " + UnitsOrdered + " " + ProductSales + " " + TotalOrderItems + "\n";
        }
    }
}
