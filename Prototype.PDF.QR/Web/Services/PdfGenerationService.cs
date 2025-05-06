using DocumentFormat.OpenXml.Packaging;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using System.Diagnostics;

namespace Web.Services;
public class PdfGenerationService
{
    private readonly IWebHostEnvironment _webHostEnvironment;

    public PdfGenerationService(IWebHostEnvironment webHostEnvironment)
    {
        _webHostEnvironment = webHostEnvironment;
    }

    public async Task<string> ConvertWordToPdfAsync(string wordFilePath)
    {
        // Generate output PDF path
        string pdfFileName = Path.GetFileNameWithoutExtension(wordFilePath) + ".pdf";
        string pdfDirectory = Path.Combine(_webHostEnvironment.WebRootPath, "outputs", "pdfs");

        // Ensure directory exists
        Directory.CreateDirectory(pdfDirectory);

        string pdfFilePath = Path.Combine(pdfDirectory, pdfFileName);

        // There are multiple approaches to convert Word to PDF:
        // 1. Use Microsoft.Office.Interop.Word (requires Word installed on server)
        // 2. Use third-party libraries like Aspose.Words (requires license)
        // 3. Use open-source libraries like DocX and iText
        // 4. Use process-based approach with LibreOffice/OpenOffice

        // For this example, we'll use a process-based approach with LibreOffice
        // assuming it's installed on the server

        try
        {
            // LibreOffice approach
            // Note: You need to have LibreOffice installed on the server
            // and the path might be different depending on the OS
            string libreOfficeExePath = @"C:\Program Files\LibreOffice\program\soffice.exe";

            // Check if LibreOffice exists
            if (!File.Exists(libreOfficeExePath))
            {
                // Fallback to basic approach using iText (limited compatibility)
                await BasicWordToPdfConversionAsync(wordFilePath, pdfFilePath);
            }
            else
            {
                // Use LibreOffice for conversion
                ProcessStartInfo psi = new ProcessStartInfo
                {
                    FileName = libreOfficeExePath,
                    Arguments = $"--headless --convert-to pdf --outdir \"{Path.GetDirectoryName(pdfFilePath)}\" \"{wordFilePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };

                using Process process = new Process { StartInfo = psi };
                process.Start();
                await process.WaitForExitAsync();

                if (process.ExitCode != 0)
                {
                    string error = await process.StandardError.ReadToEndAsync();
                    throw new Exception($"PDF conversion failed: {error}");
                }
            }

            return pdfFilePath;
        }
        catch (Exception ex)
        {
            // Log the error
            Console.WriteLine($"Error converting Word to PDF: {ex.Message}");

            // Try the basic fallback approach
            await BasicWordToPdfConversionAsync(wordFilePath, pdfFilePath);

            return pdfFilePath;
        }
    }

    private async Task BasicWordToPdfConversionAsync(string wordFilePath, string pdfFilePath)
    {
        // This is a very basic approach with limited compatibility
        // It extracts text from Word and creates a simple PDF
        // For production use, consider using a proper conversion library

        try
        {
            // Extract text from Word document
            string extractedText = ExtractTextFromWord(wordFilePath);

            // Create PDF with iText7
            await Task.Run(() =>
            {
                using var writer = new PdfWriter(pdfFilePath);
                using var pdf = new PdfDocument(writer);
                using var document = new Document(pdf);

                document.Add(new Paragraph(extractedText));
            });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in basic conversion: {ex.Message}");
            throw;
        }
    }

    private string ExtractTextFromWord(string wordFilePath)
    {
        using WordprocessingDocument doc = WordprocessingDocument.Open(wordFilePath, false);
        var body = doc.MainDocumentPart?.Document.Body;

        if (body == null)
            return string.Empty;

        return body.InnerText;
    }
}