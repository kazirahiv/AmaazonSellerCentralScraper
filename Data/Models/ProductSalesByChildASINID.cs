using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class ProductSalesByChildASINID
    {
        [Key]
        public int Id { get; set; }
        public int ChildASINId { get; set; }
        public int DateID { get; set; }
        public decimal Earning { get; set; }
    }
}
