using Microsoft.EntityFrameworkCore;
using Web.Models;

namespace Web.Data;

public class ApplicationDbContext : DbContext
{
    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options)
        : base(options)
    {
    }

    public DbSet<TemplateModel> Templates { get; set; }
    public DbSet<DocumentModel> Documents { get; set; }

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        base.OnModelCreating(modelBuilder);

        // Configure relationships
        modelBuilder.Entity<DocumentModel>()
            .HasOne(d => d.Template)
            .WithMany(t => t.Documents)
            .HasForeignKey(d => d.TemplateId);
    }
}