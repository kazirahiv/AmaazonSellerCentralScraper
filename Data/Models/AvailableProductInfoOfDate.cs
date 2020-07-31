using System;
using System.ComponentModel.DataAnnotations;

namespace Data.Models
{
    public class AvailableProductInfoOfDate
    {
        [Key]
        public int Id { get; set; }
        public DateTime DatePickerDate { get; set; }
    }
}
