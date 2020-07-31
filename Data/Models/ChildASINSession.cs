using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class ChildASINSession
    {
        [Key]
        public int Id { get; set; }
        public int ChildASINId { get; set; }
        public int DateID { get; set; }
        public int SessionValue { get; set; }
    }
}
