using Microsoft.AspNetCore.Mvc.Rendering;

namespace Web.Models.ViewModels;

public class DocumentEditorViewModel
{
    public DocumentModel Document { get; set; } = new DocumentModel();
    public SelectList? TemplateOptions { get; set; }
    public string? PreviewImagePath { get; set; }
}
