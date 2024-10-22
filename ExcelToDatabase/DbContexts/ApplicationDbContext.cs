using Microsoft.EntityFrameworkCore;
using ETS.DataImporter.DbContexts.Entities;

namespace ETS.DataImporter.DbContexts
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
        {
        }

        public DbSet<ExcelFormulaMaster> ExcelFormulaMasters { get; set; }
        public DbSet<ExcelPathMaster> ExcelPathMasters { get; set; }
        public DbSet<ExcelValues> ExcelValues { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            // ExcelFormulaMaster configuration
            modelBuilder.Entity<ExcelFormulaMaster>(entity =>
            {
                entity.ToTable("TBL_GM_EXCEL_FORMULA_MASTER");
                entity.HasKey(e => e.GM_EXCEL_FORMULA_ID);
                entity.Property(e => e.GM_EXCEL_PRODUCT_NAME)
                    .HasMaxLength(50)
                    .HasColumnType("varchar");
                entity.HasOne(e => e.ExcelPathMaster)
                    .WithMany(p => p.ExcelFormulaMasters)
                    .HasForeignKey(e => e.GM_EXCEL_PATH_MASTER_ID);
            });

            // ExcelPathMaster configuration
            modelBuilder.Entity<ExcelPathMaster>(entity =>
            {
                entity.ToTable("TBL_GM_EXCEL_PATH_MASTER");
                entity.HasKey(e => e.GMExcel_ID);
                entity.Property(e => e.GMExcel_Name).HasMaxLength(150).HasColumnType("varchar");
                entity.Property(e => e.GMExcel_Path).HasColumnType("text");
                entity.Property(e => e.GMExcel_Description).HasMaxLength(150).HasColumnType("varchar");
            });

            // ExcelValues configuration
            modelBuilder.Entity<ExcelValues>(entity =>
            {
                entity.ToTable("TBL_GM_EXCEL_VALUES");
                entity.HasKey(e => e.GM_EXCEL_VALUES_ID);
                entity.Property(e => e.GM_EXCEL_PRODUCT_NAME).HasMaxLength(150).HasColumnType("varchar");

                // Ensure GM_EXCEL_DATETIME uses 'timestamp without time zone'
                entity.Property(e => e.GM_EXCEL_DATETIME)
                      .HasColumnType("timestamp without time zone"); // Change the column type explicitly

                // Ensure GM_EXCEL_INSERTED_DATE uses 'timestamp without time zone'
                entity.Property(e => e.GM_EXCEL_INSERTED_DATE)
                      .HasColumnType("timestamp without time zone"); // Change the column type explicitly

                entity.HasOne(e => e.ExcelPathMaster)
                    .WithMany(p => p.ExcelValues)
                    .HasForeignKey(e => e.GM_EXCEL_PATH_MASTER_ID);
            });
        }
    }
}
