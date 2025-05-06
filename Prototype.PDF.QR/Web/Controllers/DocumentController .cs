using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Web.Data;
using Web.Models;
using Web.Models.ViewModels;
using Web.Services;

namespace Web.Controllers;

public class DocumentController : Controller
{
    private readonly ApplicationDbContext _context;
    private readonly IWebHostEnvironment _webHostEnvironment;
    private readonly WordProcessingService _wordProcessingService;
    private readonly PdfGenerationService _pdfGenerationService;

    public DocumentController(
        ApplicationDbContext context,
        IWebHostEnvironment webHostEnvironment,
        WordProcessingService wordProcessingService,
        PdfGenerationService pdfGenerationService)
    {
        _context = context;
        _webHostEnvironment = webHostEnvironment;
        _wordProcessingService = wordProcessingService;
        _pdfGenerationService = pdfGenerationService;
    }

    // GET: Document
    public async Task<IActionResult> Index()
    {
        var documents = await _context.Documents
            .Include(d => d.Template)
            .ToListAsync();

        return View(documents);
    }

    // GET: Document/Details/5
    public async Task<IActionResult> Details(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var document = await _context.Documents
            .Include(d => d.Template)
            .FirstOrDefaultAsync(m => m.Id == id);

        if (document == null)
        {
            return NotFound();
        }

        return View(document);
    }

    // GET: Document/Create
    public async Task<IActionResult> Create()
    {
        var templates = await _context.Templates.ToListAsync();

        if (!templates.Any())
        {
            TempData["WarningMessage"] = "You need to upload at least one template before creating a document.";
            return RedirectToAction("Upload", "Template");
        }

        var viewModel = new DocumentEditorViewModel
        {
            Document = new DocumentModel(),
            TemplateOptions = new SelectList(templates, "Id", "Name")
        };

        return View(viewModel);
    }

    // POST: Document/Create
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Create(DocumentEditorViewModel viewModel)
    {
        if (ModelState.IsValid)
        {
            try
            {
                // Get template
                var template = await _context.Templates.FindAsync(viewModel.Document.TemplateId);

                if (template == null)
                {
                    ModelState.AddModelError("Document.TemplateId", "Selected template not found.");
                    await PrepareTemplateOptions(viewModel);
                    return View(viewModel);
                }

                // Process Word document with content
                string processedWordPath = await _wordProcessingService.ProcessTemplateWithContent(
                    template.Id,
                    template.FilePath,
                    viewModel.Document.Content);

                // Convert to PDF
                string pdfPath = await _pdfGenerationService.ConvertWordToPdfAsync(processedWordPath);

                // Calculate relative paths for storage
                string relativePdfPath = Path.GetRelativePath(_webHostEnvironment.WebRootPath, pdfPath);

                // Save document record
                viewModel.Document.OutputPdfPath = relativePdfPath;
                viewModel.Document.CreatedDate = DateTime.Now;

                _context.Documents.Add(viewModel.Document);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Document created successfully.";
                return RedirectToAction(nameof(Details), new { id = viewModel.Document.Id });
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", $"Error creating document: {ex.Message}");
            }
        }

        await PrepareTemplateOptions(viewModel);
        return View(viewModel);
    }

    // GET: Document/Edit/5
    public async Task<IActionResult> Edit(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var document = await _context.Documents.FindAsync(id);

        if (document == null)
        {
            return NotFound();
        }

        var viewModel = new DocumentEditorViewModel
        {
            Document = document,
            TemplateOptions = new SelectList(await _context.Templates.ToListAsync(), "Id", "Name", document.TemplateId)
        };

        return View(viewModel);
    }

    // POST: Document/Edit/5
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Edit(int id, DocumentEditorViewModel viewModel)
    {
        if (id != viewModel.Document.Id)
        {
            return NotFound();
        }

        if (ModelState.IsValid)
        {
            try
            {
                // Get template
                var template = await _context.Templates.FindAsync(viewModel.Document.TemplateId);

                if (template == null)
                {
                    ModelState.AddModelError("Document.TemplateId", "Selected template not found.");
                    await PrepareTemplateOptions(viewModel);
                    return View(viewModel);
                }

                // Process Word document with content
                string processedWordPath = await _wordProcessingService.ProcessTemplateWithContent(
                    template.Id,
                    template.FilePath,
                    viewModel.Document.Content);

                // Convert to PDF
                string pdfPath = await _pdfGenerationService.ConvertWordToPdfAsync(processedWordPath);

                // Calculate relative paths for storage
                string relativePdfPath = Path.GetRelativePath(_webHostEnvironment.WebRootPath, pdfPath);

                // Update document record
                viewModel.Document.OutputPdfPath = relativePdfPath;
                viewModel.Document.ModifiedDate = DateTime.Now;

                _context.Update(viewModel.Document);
                await _context.SaveChangesAsync();

                TempData["SuccessMessage"] = "Document updated successfully.";
                return RedirectToAction(nameof(Details), new { id = viewModel.Document.Id });
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!DocumentExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }
            catch (Exception ex)
            {
                ModelState.AddModelError("", $"Error updating document: {ex.Message}");
            }
        }

        await PrepareTemplateOptions(viewModel);
        return View(viewModel);
    }

    // GET: Document/Delete/5
    public async Task<IActionResult> Delete(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var document = await _context.Documents
            .Include(d => d.Template)
            .FirstOrDefaultAsync(m => m.Id == id);

        if (document == null)
        {
            return NotFound();
        }

        return View(document);
    }

    // POST: Document/Delete/5
    [HttpPost, ActionName("Delete")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteConfirmed(int id)
    {
        var document = await _context.Documents.FindAsync(id);

        if (document == null)
        {
            return NotFound();
        }

        // Delete PDF file if it exists
        string pdfPath = Path.Combine(_webHostEnvironment.WebRootPath, document.OutputPdfPath ?? "");
        if (System.IO.File.Exists(pdfPath))
        {
            System.IO.File.Delete(pdfPath);
        }

        _context.Documents.Remove(document);
        await _context.SaveChangesAsync();

        TempData["SuccessMessage"] = "Document deleted successfully.";
        return RedirectToAction(nameof(Index));
    }

    // GET: Document/Preview/5
    public async Task<IActionResult> Preview(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var document = await _context.Documents
            .Include(d => d.Template)
            .FirstOrDefaultAsync(m => m.Id == id);

        if (document == null)
        {
            return NotFound();
        }

        // Check if PDF exists
        if (string.IsNullOrEmpty(document.OutputPdfPath))
        {
            TempData["ErrorMessage"] = "PDF file not found for this document.";
            return RedirectToAction(nameof(Details), new { id });
        }

        // Return PDF file
        string pdfPath = Path.Combine(_webHostEnvironment.WebRootPath, document.OutputPdfPath);
        if (!System.IO.File.Exists(pdfPath))
        {
            TempData["ErrorMessage"] = "PDF file not found on server.";
            return RedirectToAction(nameof(Details), new { id });
        }

        return File(System.IO.File.OpenRead(pdfPath), "application/pdf");
    }

    // GET: Document/Download/5
    public async Task<IActionResult> Download(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var document = await _context.Documents
            .Include(d => d.Template)
            .FirstOrDefaultAsync(m => m.Id == id);

        if (document == null)
        {
            return NotFound();
        }

        // Check if PDF exists
        if (string.IsNullOrEmpty(document.OutputPdfPath))
        {
            TempData["ErrorMessage"] = "PDF file not found for this document.";
            return RedirectToAction(nameof(Details), new { id });
        }

        // Return PDF file for download
        string pdfPath = Path.Combine(_webHostEnvironment.WebRootPath, document.OutputPdfPath);
        if (!System.IO.File.Exists(pdfPath))
        {
            TempData["ErrorMessage"] = "PDF file not found on server.";
            return RedirectToAction(nameof(Details), new { id });
        }

        string fileName = $"{document.Title}_{DateTime.Now:yyyyMMdd}.pdf";
        return PhysicalFile(pdfPath, "application/pdf", fileName);
    }

    private async Task PrepareTemplateOptions(DocumentEditorViewModel viewModel)
    {
        viewModel.TemplateOptions = new SelectList(
            await _context.Templates.ToListAsync(),
            "Id",
            "Name",
            viewModel.Document.TemplateId);
    }

    private bool DocumentExists(int id)
    {
        return _context.Documents.Any(e => e.Id == id);
    }
}