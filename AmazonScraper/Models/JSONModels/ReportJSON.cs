namespace AmazonScraper.Models.JSONModels
{
    public class ReportJSON
    {
        public string requestId { get; set; }
        public Headers headers { get; set; }
        public CacheForm cacheForm { get; set; }
        public Data data { get; set; }
        public int isTimeout { get; set; }
        public int isDataError { get; set; }
        public int isOldData { get; set; }
        public Graph graph { get; set; }
        public int isReportDefinitionError { get; set; }
    }
}
