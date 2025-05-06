using System.ComponentModel.DataAnnotations;

namespace Web.Models;

public class TemplateModel
{
    public int Id { get; set; }

    [Required]
    [StringLength(100)]
    public string Name { get; set; } = string.Empty;

    [StringLength(255)]
    public string Description { get; set; } = string.Empty;

    [Required]
    [StringLength(255)]
    public string FilePath { get; set; } = string.Empty;

    [Required]
    [StringLength(255)]
    public string OriginalFileName { get; set; } = string.Empty;

    public DateTime UploadedDate { get; set; } = DateTime.Now;

    // Navigation property
    public virtual ICollection<DocumentModel>? Documents { get; set; }
}
