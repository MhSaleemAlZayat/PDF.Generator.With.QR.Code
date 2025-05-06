using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Web.Data;
using Web.Models;
using Web.Models.ViewModels;

namespace Web.Controllers;

public class TemplateController : Controller
{
    private readonly ApplicationDbContext _context;
    private readonly IWebHostEnvironment _webHostEnvironment;
    private readonly IConfiguration _configuration;

    public TemplateController(
        ApplicationDbContext context,
        IWebHostEnvironment webHostEnvironment,
        IConfiguration configuration)
    {
        _context = context;
        _webHostEnvironment = webHostEnvironment;
        _configuration = configuration;
    }

    // GET: Template
    public async Task<IActionResult> Index()
    {
        var templates = await _context.Templates.ToListAsync();
        return View(templates);
    }

    // GET: Template/Details/5
    public async Task<IActionResult> Details(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var template = await _context.Templates
            .FirstOrDefaultAsync(m => m.Id == id);

        if (template == null)
        {
            return NotFound();
        }

        return View(template);
    }

    // GET: Template/Upload
    public IActionResult Upload()
    {
        return View();
    }

    // POST: Template/Upload
    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> Upload(TemplateUploadViewModel model)
    {
        if (ModelState.IsValid)
        {
            if (model.TemplateFile != null && model.TemplateFile.Length > 0)
            {
                // Validate file extension (allow only .docx)
                string extension = Path.GetExtension(model.TemplateFile.FileName).ToLowerInvariant();
                if (extension != ".docx")
                {
                    ModelState.AddModelError("TemplateFile", "Only Word documents (.docx) are allowed.");
                    return View(model);
                }

                try
                {
                    // Generate unique filename
                    string uniqueFileName = $"{Guid.NewGuid()}{extension}";
                    string uploadsFolder = Path.Combine(_webHostEnvironment.WebRootPath, "templates");

                    // Ensure directory exists
                    Directory.CreateDirectory(uploadsFolder);

                    string filePath = Path.Combine(uploadsFolder, uniqueFileName);

                    // Save the file
                    using (var fileStream = new FileStream(filePath, FileMode.Create))
                    {
                        await model.TemplateFile.CopyToAsync(fileStream);
                    }

                    // Create template record
                    var template = new TemplateModel
                    {
                        Name = model.Name,
                        Description = model.Description,
                        FilePath = filePath,
                        OriginalFileName = model.TemplateFile.FileName,
                        UploadedDate = DateTime.Now
                    };

                    _context.Templates.Add(template);
                    await _context.SaveChangesAsync();

                    TempData["SuccessMessage"] = "Template uploaded successfully.";
                    return RedirectToAction(nameof(Index));
                }
                catch (Exception ex)
                {
                    ModelState.AddModelError("", $"Error uploading template: {ex.Message}");
                }
            }
            else
            {
                ModelState.AddModelError("TemplateFile", "Please select a file to upload.");
            }
        }

        return View(model);
    }

    // GET: Template/Delete/5
    public async Task<IActionResult> Delete(int? id)
    {
        if (id == null)
        {
            return NotFound();
        }

        var template = await _context.Templates
            .FirstOrDefaultAsync(m => m.Id == id);

        if (template == null)
        {
            return NotFound();
        }

        return View(template);
    }

    // POST: Template/Delete/5
    [HttpPost, ActionName("Delete")]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> DeleteConfirmed(int id)
    {
        var template = await _context.Templates.FindAsync(id);

        if (template == null)
        {
            return NotFound();
        }

        // Delete the file if it exists
        if (System.IO.File.Exists(template.FilePath))
        {
            System.IO.File.Delete(template.FilePath);
        }

        _context.Templates.Remove(template);
        await _context.SaveChangesAsync();

        TempData["SuccessMessage"] = "Template deleted successfully.";
        return RedirectToAction(nameof(Index));
    }
}
