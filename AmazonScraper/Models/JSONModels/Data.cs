using System.Collections.Generic;

namespace AmazonScraper.Models.JSONModels
{
    public class Data
    {
        public string reportDesc { get; set; }
        public string reportTitle { get; set; }
        public int hasNextPage { get; set; }
        public List<List<string>> rows { get; set; }
    }
}
