using Microsoft.EntityFrameworkCore;
using SowAutomationTool.Models;

namespace SowAutomationTool.Data
{
    public class AppDbContext : DbContext
    {
        public AppDbContext(DbContextOptions<AppDbContext> options) : base(options) { }

        public DbSet<AppUser> Users => Set<AppUser>();

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<AppUser>()
                .HasIndex(u => u.Email)
                .IsUnique();

            // --- Seed new user here ---
            modelBuilder.Entity<AppUser>().HasData(
                new AppUser
                {
                    Id = 4, // Must be provided for seeding
                    Email = "christo.kl@gds.ey.com",
                    DisplayName = "christo",
                    Role = "Admin",
                    IsActive = true
                }
            );
        }
    }
}



