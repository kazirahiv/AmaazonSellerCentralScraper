namespace AmazonScraper.Models.JSONModels
{
    public class CacheForm
    {
        public int sortColumn { get; set; }
        public string realID { get; set; }
        public int dateUnit { get; set; }
        public string reportID { get; set; }
        public int currentPage { get; set; }
        public int sortIsAscending { get; set; }
        public int defaultDatesUsed { get; set; }
        public string viewDateUnits { get; set; }
        public string fromDate { get; set; }
        public string toDate { get; set; }
    }
}
