using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace PdfImageExtractor;

class Program
{
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "This application only supports Windows.")]
    static void Main(string[] args)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Console.WriteLine("This application only runs on Windows.");
            return;
        }

        var supportedFormats = typeof (ImageFormat).GetProperties(BindingFlags.Public | BindingFlags.Static).Where(p => p.PropertyType == typeof(ImageFormat)).OrderBy(p => p.Name).ToList();

        string? pdfPath = args.ElementAtOrDefault(0);
        string? outputDirectory = args.ElementAtOrDefault(1);

        if (string.IsNullOrWhiteSpace(pdfPath) || string.IsNullOrWhiteSpace(outputDirectory))
        {
            Console.WriteLine("");
            Console.WriteLine("Usage: PdfImageExtractor <pathToPDF> <pathToFolderForImages> <imageFormat>");
            Console.WriteLine("");
            Console.WriteLine("The default imageFormat is Png.");
            Console.WriteLine("Supported values are " + string.Join(", ", supportedFormats.Select(p => p.Name)));
            Console.WriteLine("");
            return;
        }

        if (!File.Exists(pdfPath))
        {
            Console.WriteLine("The specified pathToPDF does not exist");
            return;
        }

        if (!Directory.Exists(outputDirectory))
        {
            Console.WriteLine("The specified pathToFolderForImages does not exist");
            return;
        }

        var imageFormat = ImageFormat.Png;
        var imageFormatName = (args.ElementAtOrDefault(2) ?? "Png").Trim().ToLowerInvariant();
        var requestedFormat = supportedFormats
            .Where(p => p.Name.Equals(imageFormatName, StringComparison.OrdinalIgnoreCase))
            .Select(p => p.GetValue(null) as ImageFormat)
            .FirstOrDefault();
        if (requestedFormat is not null)
        {
            imageFormat = requestedFormat;
        }

        PdfDocument pdfDoc = new(new PdfReader(pdfPath));
        var numPages = pdfDoc.GetNumberOfPages();
        var numDigits = numPages.ToString().Length;
        for (int pageNumber = 1; pageNumber <= numPages; pageNumber++)
        {
            var imageCount = 0;
            var page = pdfDoc.GetPage(pageNumber);
            var imageRenderListener = new ImageRenderListener(outputDirectory, imageFormat, () =>
            {
                imageCount++;
                return $"Page-{pageNumber.ToString().PadLeft(numDigits, '0')}-Image-{imageCount.ToString().PadLeft(3, '0')}.{imageFormatName}";
            });
            var parser = new PdfCanvasProcessor(imageRenderListener);
            parser.ProcessPageContent(page);
        }

        pdfDoc.Close();
        Console.WriteLine("Images extracted.");
    }
}

[System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "This application only supports Windows.")]
internal class ImageRenderListener(string outputDir, ImageFormat format, Func<string> getImageName) : IEventListener
{
    private readonly string outputDir = outputDir;

    public void EventOccurred(IEventData data, EventType type)
    {
        if (type == EventType.RENDER_IMAGE && data is ImageRenderInfo renderInfo)
        {
            var imageObject = renderInfo.GetImage();
            var imageBytes = imageObject.GetImageBytes();

            using var imageStream = new MemoryStream(imageBytes);
            using var image = Image.FromStream(imageStream);
            string filePath = Path.Combine(outputDir, getImageName());
            image.Save(filePath, format);
        }
    }

    public ICollection<EventType> GetSupportedEvents()
    {
        return [EventType.RENDER_IMAGE];
    }
}