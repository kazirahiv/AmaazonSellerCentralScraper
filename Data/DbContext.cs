using Data.Models;
using Microsoft.EntityFrameworkCore;

namespace Data
{
    public class AmazonDBContext : DbContext
    {
        public virtual DbSet<UniqueProductASIN> UniqueProductASINs { get; set; }
        public virtual DbSet<AvailableProductInfoOfDate> AvailableProductInfoOfDates { get; set; }
        public virtual DbSet<ChildASINSession> ChildASINSessions { get; set; }
        public virtual DbSet<UnitsOrderedByASINID> UnitsOrderedByASINIDs { get; set; }
        public virtual DbSet<ProductSalesByChildASINID> ProductSalesByChildASINIDs { get; set; }
        public virtual DbSet<TotalOrderItemsByASINID> TotalOrderItemsByASINIDs { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(@"Server=.\SQLEXPRESS;Database=AmazonScraperDB;Trusted_Connection=True;");
        }
    }
}
