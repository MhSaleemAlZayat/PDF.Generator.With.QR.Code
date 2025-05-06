using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;

namespace Web.Services;

public class WordProcessingService
{
    private readonly IWebHostEnvironment _webHostEnvironment;
    private readonly IConfiguration _configuration;
    private readonly QrCodeService _qrCodeService;

    public WordProcessingService(
        IWebHostEnvironment webHostEnvironment,
        IConfiguration configuration,
        QrCodeService qrCodeService)
    {
        _webHostEnvironment = webHostEnvironment;
        _configuration = configuration;
        _qrCodeService = qrCodeService;
    }

    public async Task<string> ProcessTemplateWithContent(int templateId, string templatePath, string content)
    {
        // Generate a unique filename for the processed document
        string outputFileName = $"processed_{templateId}_{DateTime.Now:yyyyMMddHHmmss}.docx";
        string outputPath = Path.Combine(_webHostEnvironment.WebRootPath, "outputs", outputFileName);

        // Make sure the output directory exists
        Directory.CreateDirectory(Path.GetDirectoryName(outputPath)!);

        // Copy the template to the output path
        File.Copy(templatePath, outputPath, true);

        // Extract URLs from content for QR code generation
        var urls = ExtractUrls(content);
        var qrCodePaths = new Dictionary<string, string>();

        // Generate QR codes for each URL
        foreach (var url in urls)
        {
            var qrCodePath = await _qrCodeService.GenerateQrCodeAsync(url);
            qrCodePaths.Add(url, qrCodePath);
        }

        // Open the document
        using (WordprocessingDocument doc = WordprocessingDocument.Open(outputPath, true))
        {
            MainDocumentPart mainPart = doc.MainDocumentPart!;

            // Find a placeholder in the document to insert the content
            // This is a simple implementation - you might want to enhance this
            // to support multiple placeholders or specific formatting
            Body body = mainPart.Document.Body!;

            // Find paragraph with placeholder text
            var paragraphs = body.Elements<Paragraph>().ToList();

            // Look for placeholder {{CONTENT}}
            Paragraph? contentPlaceholder = null;
            foreach (var paragraph in paragraphs)
            {
                if (paragraph.InnerText.Contains("{{CONTENT}}"))
                {
                    contentPlaceholder = paragraph;
                    break;
                }
            }

            if (contentPlaceholder != null)
            {
                // Replace the placeholder with the actual content
                // This is a simple replacement - in a real application,
                // you might want to use OpenXML more extensively to
                // preserve formatting or handle more complex cases
                var runs = contentPlaceholder.Elements<Run>().ToList();
                foreach (var run in runs)
                {
                    string runText = run.InnerText;
                    if (runText.Contains("{{CONTENT}}"))
                    {
                        run.RemoveAllChildren<Text>();
                        run.AppendChild(new Text(runText.Replace("{{CONTENT}}", content)));
                    }
                }
            }
            else
            {
                // If no placeholder is found, append content to the end
                var newParagraph = new Paragraph(new Run(new Text(content)));
                body.AppendChild(newParagraph);
            }

            // Insert QR codes as images
            // For each URL, insert a QR code image after the content
            if (qrCodePaths.Any())
            {
                // Insert a new paragraph for QR codes
                var qrCodeParagraph = new Paragraph(new Run(new Text("QR Codes for embedded URLs:")));
                body.AppendChild(qrCodeParagraph);

                foreach (var qrCode in qrCodePaths)
                {
                    // Insert QR code image
                    var imageId = InsertImage(mainPart, qrCode.Value);

                    // Create a paragraph for the QR code with URL description
                    var urlParagraph = new Paragraph(
                        new Run(
                            new Text($"QR Code for: {qrCode.Key}"),
                            new Break(),
                            new Drawing(
                                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = 1000000, Cy = 1000000 },
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties { Id = 1, Name = "QR Code" },
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(),
                                    new DocumentFormat.OpenXml.Drawing.Graphic(
                                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties { Id = 0, Name = "QR Code" },
                                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                                ),
                                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                                    new DocumentFormat.OpenXml.Drawing.Blip { Embed = imageId },
                                                    new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())
                                                ),
                                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                                        new DocumentFormat.OpenXml.Drawing.Offset { X = 0, Y = 0 },
                                                        new DocumentFormat.OpenXml.Drawing.Extents { Cx = 1000000, Cy = 1000000 }
                                                    ),
                                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
                                                )
                                            )
                                        )
                                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                                    )
                                )
                                { DistanceFromTop = 0, DistanceFromBottom = 0, DistanceFromLeft = 0, DistanceFromRight = 0 }
                            )
                        )
                    );

                    body.AppendChild(urlParagraph);
                }
            }

            // Save the changes
            mainPart.Document.Save();
        }

        return outputPath;
    }

    private string InsertImage(MainDocumentPart mainPart, string imagePath)
    {
        var imagePartType = GetImagePartType(imagePath);
        ImagePart imagePart = mainPart.AddImagePart(imagePartType);

        using (FileStream stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        return mainPart.GetIdOfPart(imagePart);
    }

    private ImagePartType GetImagePartType(string imagePath)
    {
        string extension = Path.GetExtension(imagePath).ToLower();
        return extension switch
        {
            ".png" => ImagePartType.Png,
            ".jpg" or ".jpeg" => ImagePartType.Jpeg,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            _ => ImagePartType.Png // Default to PNG
        };
    }

    private List<string> ExtractUrls(string content)
    {
        var urls = new List<string>();

        // Regular expression to match URLs
        var regex = new Regex(@"https?:\/\/(www\.)?[-a-zA-Z0-9@:%._\+~#=]{1,256}\.[a-zA-Z0-9()]{1,6}\b([-a-zA-Z0-9()@:%_\+.~#?&//=]*)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        var matches = regex.Matches(content);
        foreach (Match match in matches)
        {
            urls.Add(match.Value);
        }

        return urls;
    }
}



