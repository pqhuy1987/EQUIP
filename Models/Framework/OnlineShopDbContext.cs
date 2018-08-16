namespace Models.Framework
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class OnlineShopDbContext : DbContext
    {
        public OnlineShopDbContext()
            : base("name=OnlineShopDbContext")
        {
        }

        public virtual DbSet<Account> Accounts { get; set; }
        public virtual DbSet<Catelory> Catelories { get; set; }
        public virtual DbSet<LLTC> LLTCs { get; set; }
        public virtual DbSet<Product> Products { get; set; }
        public virtual DbSet<Thiet_Bi> Thiet_Bis { get; set; }
        public virtual DbSet<WorkCount> WorkCounts { get; set; }
        public virtual DbSet<CS_tbPhong_Ban> CS_tbPhong_Ban { get; set; }
        public virtual DbSet<CS_tbWorkType> CS_tbWorkType { get; set; }
        public virtual DbSet<CS_tbViTri> CS_tbViTri { get; set; }
        public virtual DbSet<CS_tbLLTCTypeSub> CS_tbLLTCTypeSub { get; set; }
        public virtual DbSet<CS_tbWorkCount> CS_tbWorkCount { get; set; }
        public virtual DbSet<CS_tbWorkCount_Sub> CS_tbWorkCount_Sub { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Catelory>()
                .Property(e => e.Email)
                .IsFixedLength();

            modelBuilder.Entity<Product>()
                .Property(e => e.Alias)
                .IsUnicode(false);

            modelBuilder.Entity<Product>()
                .Property(e => e.Price)
                .HasPrecision(18, 0);
        }
    }
}
