using System.ComponentModel.DataAnnotations;

namespace Web.Models.ViewModels;

public class TemplateUploadViewModel
{
    [Required]
    [StringLength(100)]
    public string Name { get; set; } = string.Empty;

    [StringLength(255)]
    public string Description { get; set; } = string.Empty;

    [Required(ErrorMessage = "Please select a Word document template")]
    public IFormFile? TemplateFile { get; set; }
}
