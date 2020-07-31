using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class UnitsOrderedByASINID
    {
        [Key]
        public int Id { get; set; }
        public int ChildASINId { get; set; }
        public int DateID { get; set; }
        public int UnitsOrdered { get; set; }
    }
}
