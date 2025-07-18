using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

namespace PRB.Repository.DataContext
{
    public partial class PRB_DB_Context : DbContext
    {
        public PRB_DB_Context()
        {
        }

        public PRB_DB_Context(DbContextOptions<PRB_DB_Context> options)
            : base(options)
        {
        }

        public virtual DbSet<PrbCompanyPrice> PrbCompanyPrices { get; set; } = null!;
        public virtual DbSet<PrbCurrencyCode> PrbCurrencyCodes { get; set; } = null!;
        public virtual DbSet<PrbHoldingDetail> PrbHoldingDetails { get; set; } = null!;
        public virtual DbSet<PrbReportStatus> PrbReportStatuses { get; set; } = null!;
        public virtual DbSet<PrbReportStatusCode> PrbReportStatusCodes { get; set; } = null!;
        public virtual DbSet<PrbReportSummary> PrbReportSummaries { get; set; } = null!;
        public virtual DbSet<PrbReportType> PrbReportTypes { get; set; } = null!;
        public virtual DbSet<PrbReportTypeCode> PrbReportTypeCodes { get; set; } = null!;
        public virtual DbSet<PrbRoleCode> PrbRoleCodes { get; set; } = null!;
        public virtual DbSet<PrbSectorCode> PrbSectorCodes { get; set; } = null!;
        public virtual DbSet<PrbTemplatePath> PrbTemplatePaths { get; set; } = null!;
        public virtual DbSet<PrbTicker> PrbTickers { get; set; } = null!;
        public virtual DbSet<PrbTransactionTypeCode> PrbTransactionTypeCodes { get; set; } = null!;
        public virtual DbSet<PrbUser> PrbUsers { get; set; } = null!;

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see http://go.microsoft.com/fwlink/?LinkId=723263.
                optionsBuilder.UseSqlServer("Data Source=arga-az-sql02; Initial Catalog=ATM_INT2023;Trusted_connection=true; pooling=true; Max Pool Size=1000; MultipleActiveResultSets=True; Integrated Security=SSPI; TrustServerCertificate=True;Connect Timeout=120;");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<PrbCompanyPrice>(entity =>
            {
                entity.HasKey(e => new { e.CompanyTicker, e.ReportDate })
                    .HasName("PK_Company_Price");

                entity.ToTable("PRB_Company_Price", "ARGA\\vmariappan");

                entity.Property(e => e.CompanyTicker)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("Company_Ticker")
                    .IsFixedLength();

                entity.Property(e => e.ReportDate)
                    .HasColumnType("date")
                    .HasColumnName("Report_Date");

                entity.Property(e => e.LastMarketPrice)
                    .HasColumnType("decimal(7, 2)")
                    .HasColumnName("Last_Market_Price");

                entity.HasOne(d => d.CompanyTickerNavigation)
                    .WithMany(p => p.PrbCompanyPrices)
                    .HasForeignKey(d => d.CompanyTicker)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Ticker");
            });

            modelBuilder.Entity<PrbCurrencyCode>(entity =>
            {
                entity.HasKey(e => e.CurrencyCode)
                    .HasName("PK__PRB_Curr__89A7C39E4D8F1464");

                entity.ToTable("PRB_Currency_Code", "ARGA\\vmariappan");

                entity.Property(e => e.CurrencyCode)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("Currency_Code")
                    .IsFixedLength();

                entity.Property(e => e.CurrencyDesc)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("Currency_Desc");
            });

            modelBuilder.Entity<PrbHoldingDetail>(entity =>
            {
                entity.HasKey(e => new { e.CompanyTicker, e.TransactionDate })
                    .HasName("PK_Holdings_Key");

                entity.ToTable("PRB_Holding_Details", "ARGA\\vmariappan");

                entity.Property(e => e.CompanyTicker)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("Company_Ticker")
                    .IsFixedLength();

                entity.Property(e => e.TransactionDate)
                    .HasColumnType("datetime")
                    .HasColumnName("Transaction_Date");

                entity.Property(e => e.Amount).HasColumnType("decimal(7, 2)");

                entity.Property(e => e.CurrencyCode)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("Currency_Code")
                    .IsFixedLength();

                entity.Property(e => e.TransactionTypeCode)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("Transaction_Type_Code")
                    .IsFixedLength();

                entity.HasOne(d => d.CompanyTickerNavigation)
                    .WithMany(p => p.PrbHoldingDetails)
                    .HasForeignKey(d => d.CompanyTicker)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Company_Ticker");

                entity.HasOne(d => d.CurrencyCodeNavigation)
                    .WithMany(p => p.PrbHoldingDetails)
                    .HasForeignKey(d => d.CurrencyCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Currency_Code");

                entity.HasOne(d => d.TransactionTypeCodeNavigation)
                    .WithMany(p => p.PrbHoldingDetails)
                    .HasForeignKey(d => d.TransactionTypeCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Transaction_Type_Code");
            });

            modelBuilder.Entity<PrbReportStatus>(entity =>
            {
                entity.HasKey(e => new { e.ReportMonth, e.ReportDate, e.RoleCode })
                    .HasName("PK_Report_Status");

                entity.ToTable("PRB_Report_Status", "ARGA\\vmariappan");

                entity.Property(e => e.ReportMonth)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("Report_Month");

                entity.Property(e => e.ReportDate)
                    .HasColumnType("datetime")
                    .HasColumnName("Report_Date");

                entity.Property(e => e.RoleCode)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("Role_Code")
                    .IsFixedLength();

                entity.Property(e => e.Comments)
                    .HasMaxLength(250)
                    .IsUnicode(false);

                entity.Property(e => e.ReportStatusCode)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("Report_Status_Code")
                    .IsFixedLength();

                entity.HasOne(d => d.ReportStatusCodeNavigation)
                    .WithMany(p => p.PrbReportStatuses)
                    .HasForeignKey(d => d.ReportStatusCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Report_Status_Code");

                entity.HasOne(d => d.RoleCodeNavigation)
                    .WithMany(p => p.PrbReportStatuses)
                    .HasForeignKey(d => d.RoleCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Rolee_Code");
            });

            modelBuilder.Entity<PrbReportStatusCode>(entity =>
            {
                entity.HasKey(e => e.ReportStatusCode)
                    .HasName("PK__PRB_Repo__5A6F6A9234C39DD0");

                entity.ToTable("PRB_Report_Status_Code", "ARGA\\vmariappan");

                entity.Property(e => e.ReportStatusCode)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("Report_Status_Code")
                    .IsFixedLength();

                entity.Property(e => e.ReportStatusDesc)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("Report_Status_Desc");
            });

            modelBuilder.Entity<PrbReportSummary>(entity =>
            {
                entity.HasKey(e => e.ReportMonth)
                    .HasName("PK_Report_Summary");

                entity.ToTable("PRB_Report_Summary", "ARGA\\vmariappan");

                entity.Property(e => e.ReportMonth)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("Report_Month");

                entity.Property(e => e.FilePassword)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("File_Password");

                entity.Property(e => e.ReportStatusCode)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("Report_Status_Code")
                    .IsFixedLength();

                entity.HasOne(d => d.ReportStatusCodeNavigation)
                    .WithMany(p => p.PrbReportSummaries)
                    .HasForeignKey(d => d.ReportStatusCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Reportt_Status_Code");
            });

            modelBuilder.Entity<PrbReportType>(entity =>
            {
                entity.HasKey(e => new { e.Month, e.ReportTypeCode })
                    .HasName("PK_Report_Key");

                entity.ToTable("PRB_Report_Type", "ARGA\\vmariappan");

                entity.Property(e => e.Month)
                    .HasMaxLength(10)
                    .IsUnicode(false);

                entity.Property(e => e.ReportTypeCode)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("Report_Type_Code")
                    .IsFixedLength();

                entity.Property(e => e.Commentary)
                    .HasMaxLength(2500)
                    .IsUnicode(false);

                entity.HasOne(d => d.ReportTypeCodeNavigation)
                    .WithMany(p => p.PrbReportTypes)
                    .HasForeignKey(d => d.ReportTypeCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Report_Type_Code");
            });

            modelBuilder.Entity<PrbReportTypeCode>(entity =>
            {
                entity.HasKey(e => e.ReportTypeCode)
                    .HasName("PK__PRB_Repo__47D32E08E69392B4");

                entity.ToTable("PRB_Report_Type_Code", "ARGA\\vmariappan");

                entity.Property(e => e.ReportTypeCode)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("Report_Type_Code")
                    .IsFixedLength();

                entity.Property(e => e.ReportTypeDesc)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("Report_Type_Desc");
            });

            modelBuilder.Entity<PrbRoleCode>(entity =>
            {
                entity.HasKey(e => e.RoleCode)
                    .HasName("PK__PRB_Role__1E8351060A23F890");

                entity.ToTable("PRB_Role_Code", "ARGA\\vmariappan");

                entity.Property(e => e.RoleCode)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("Role_Code")
                    .IsFixedLength();

                entity.Property(e => e.RoleDesc)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("Role_Desc");
            });

            modelBuilder.Entity<PrbSectorCode>(entity =>
            {
                entity.HasKey(e => e.SectorCode)
                    .HasName("PK__PRB_Sect__BC49ACBCE1F7A532");

                entity.ToTable("PRB_Sector_Code", "ARGA\\vmariappan");

                entity.Property(e => e.SectorCode)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("Sector_Code")
                    .IsFixedLength();

                entity.Property(e => e.SectorName)
                    .HasMaxLength(40)
                    .IsUnicode(false)
                    .HasColumnName("Sector_Name");
            });

            modelBuilder.Entity<PrbTemplatePath>(entity =>
            {
                entity.HasKey(e => e.TemplateId)
                    .HasName("Template_Id");

                entity.ToTable("PRB_Template_Path", "ARGA\\dpriscilla");

                entity.Property(e => e.TemplateId)
                    .ValueGeneratedNever()
                    .HasColumnName("Template_Id");

                entity.Property(e => e.ExpiryDate)
                    .HasColumnType("date")
                    .HasColumnName("Expiry_Date")
                    .HasDefaultValueSql("('2023-05-13')");

                entity.Property(e => e.FileName)
                    .HasMaxLength(25)
                    .IsUnicode(false)
                    .HasColumnName("File_Name");

                entity.Property(e => e.FilePath)
                    .HasMaxLength(80)
                    .IsUnicode(false)
                    .HasColumnName("File_Path");
            });

            modelBuilder.Entity<PrbTicker>(entity =>
            {
                entity.HasKey(e => e.CompanyTicker)
                    .HasName("PK__PRB_Tick__8D994BFD11F17941");

                entity.ToTable("PRB_Ticker", "ARGA\\vmariappan");

                entity.Property(e => e.CompanyTicker)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("Company_Ticker")
                    .IsFixedLength();

                entity.Property(e => e.CompanyDesc)
                    .HasMaxLength(250)
                    .IsUnicode(false)
                    .HasColumnName("Company_Desc");

                entity.Property(e => e.CompanyName)
                    .HasMaxLength(100)
                    .IsUnicode(false)
                    .HasColumnName("Company_Name");

                entity.Property(e => e.SectorCode)
                    .HasMaxLength(3)
                    .IsUnicode(false)
                    .HasColumnName("Sector_Code")
                    .IsFixedLength();

                entity.Property(e => e.TemplateId).HasColumnName("Template_Id");

                entity.HasOne(d => d.SectorCodeNavigation)
                    .WithMany(p => p.PrbTickers)
                    .HasForeignKey(d => d.SectorCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Sector_Code");

                entity.HasOne(d => d.Template)
                    .WithMany(p => p.PrbTickers)
                    .HasForeignKey(d => d.TemplateId)
                    .HasConstraintName("Template_Id");
            });

            modelBuilder.Entity<PrbTransactionTypeCode>(entity =>
            {
                entity.HasKey(e => e.TransactionTypeCode)
                    .HasName("PK__PRB_Tran__BBA313E2E6707A4D");

                entity.ToTable("PRB_Transaction_Type_Code", "ARGA\\vmariappan");

                entity.Property(e => e.TransactionTypeCode)
                    .HasMaxLength(1)
                    .IsUnicode(false)
                    .HasColumnName("Transaction_Type_Code")
                    .IsFixedLength();

                entity.Property(e => e.TransactionDesc)
                    .HasMaxLength(4)
                    .IsUnicode(false)
                    .HasColumnName("Transaction_Desc");
            });

            modelBuilder.Entity<PrbUser>(entity =>
            {
                entity.HasKey(e => e.UserMailId)
                    .HasName("PK__PRB_User__9206542B7DB8C441");

                entity.ToTable("PRB_Users", "ARGA\\vmariappan");

                entity.Property(e => e.UserMailId)
                    .HasMaxLength(100)
                    .IsUnicode(false)
                    .HasColumnName("User_MailID");

                entity.Property(e => e.Password)
                    .HasMaxLength(20)
                    .IsUnicode(false);

                entity.Property(e => e.RoleCode)
                    .HasMaxLength(2)
                    .IsUnicode(false)
                    .HasColumnName("Role_Code")
                    .IsFixedLength();

                entity.Property(e => e.UserName)
                    .HasMaxLength(200)
                    .IsUnicode(false)
                    .HasColumnName("User_Name");

                entity.HasOne(d => d.RoleCodeNavigation)
                    .WithMany(p => p.PrbUsers)
                    .HasForeignKey(d => d.RoleCode)
                    .OnDelete(DeleteBehavior.ClientSetNull)
                    .HasConstraintName("FK_Role_Code");
            });

            OnModelCreatingPartial(modelBuilder);
        }

        partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
    }
}
