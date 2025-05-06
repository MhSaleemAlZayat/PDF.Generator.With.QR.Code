using System.ComponentModel.DataAnnotations;

namespace Web.Models;

public class DocumentModel
{
    public int Id { get; set; }

    [Required]
    [StringLength(100)]
    public string Title { get; set; } = string.Empty;

    [StringLength(255)]
    public string Description { get; set; } = string.Empty;

    [Required]
    public string Content { get; set; } = string.Empty;

    public int TemplateId { get; set; }

    public string? OutputPdfPath { get; set; }

    public DateTime CreatedDate { get; set; } = DateTime.Now;
    public DateTime? ModifiedDate { get; set; }

    // Navigation property
    public virtual TemplateModel? Template { get; set; }
}
