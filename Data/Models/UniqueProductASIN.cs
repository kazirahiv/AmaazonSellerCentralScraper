using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class UniqueProductASIN
    {
        [Key]
        public int Id { get; set; }
        public string ChildAsinID { get; set; }
    }
}
