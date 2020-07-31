using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class TotalOrderItemsByASINID
    {
        [Key]
        public int Id { get; set; }
        public int ChildASINId { get; set; }
        public int DateID { get; set; }
        public decimal TotalOrders { get; set; }
    }
}
