using QRCoder;
using System.Drawing;
using System.Drawing.Imaging;


namespace Web.Services;
public class QrCodeService
{
    private readonly IWebHostEnvironment _webHostEnvironment;

    public QrCodeService(IWebHostEnvironment webHostEnvironment)
    {
        _webHostEnvironment = webHostEnvironment;
    }

    public async Task<string> GenerateQrCodeAsync(string text)
    {
        //using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
        //using (QRCodeData qrCodeData = qrGenerator.CreateQrCode("The text which should be encoded.", QRCodeGenerator.ECCLevel.Q))
        //using (PngByteQRCode qrCode = new PngByteQRCode(qrCodeData))
        //{
        //    byte[] qrCodeImage = qrCode.GetGraphic(20);
        //}

        // Create QR Code generator
        QRCodeGenerator qrGenerator = new QRCodeGenerator();
        QRCodeData qrCodeData = qrGenerator.CreateQrCode("The text which should be encoded.", QRCodeGenerator.ECCLevel.Q);
        PngByteQRCode qrCode = new PngByteQRCode(qrCodeData);
        byte[] qrCodeImageByte = qrCode.GetGraphic(20);
        Stream stream = new MemoryStream(qrCodeImageByte);
        Bitmap qrCodeImage = new Bitmap(stream);

        // Create a unique filename
        string fileName = $"qrcode_{Guid.NewGuid()}.png";
        string directory = Path.Combine(_webHostEnvironment.WebRootPath, "outputs", "qrcodes");

        // Ensure directory exists
        Directory.CreateDirectory(directory);

        string filePath = Path.Combine(directory, fileName);

        // Save the QR code image
        await Task.Run(() => qrCodeImage.Save(filePath, ImageFormat.Png));


        return filePath;
    }
}