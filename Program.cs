using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace PdfImageExtractor;

class Program
{
    static void Main(string[] args)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            Console.WriteLine("This application only runs on Windows.");
            return;
        }

        string? pdfPath = args.ElementAtOrDefault(0);
        string? outputDirectory = args.ElementAtOrDefault(1);

        if (string.IsNullOrWhiteSpace(pdfPath) || string.IsNullOrWhiteSpace(outputDirectory))
        {
            Console.WriteLine("Usage: PdfImageExtractor <pathToPDF> <pathToFolderForImages>");
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

        PdfDocument pdfDoc = new(new PdfReader(pdfPath));

        for (int pageNumber = 1; pageNumber <= pdfDoc.GetNumberOfPages(); pageNumber++)
        {
            var page = pdfDoc.GetPage(pageNumber);
            var strategy = new ImageRenderListener(outputDirectory);
            PdfCanvasProcessor parser = new(strategy);
            parser.ProcessPageContent(page);
        }

        pdfDoc.Close();
        Console.WriteLine("Images extracted.");
    }
}

[System.Diagnostics.CodeAnalysis.SuppressMessage("Interoperability", "CA1416:Validate platform compatibility", Justification = "This application only supports Windows.")]
internal class ImageRenderListener(string outputDir) : IEventListener
{
    private readonly string outputDir = outputDir;

    public void EventOccurred(IEventData data, EventType type)
    {
        if (type == EventType.RENDER_IMAGE && data is ImageRenderInfo renderInfo)
        {
            var imageObject = renderInfo.GetImage();
            var imageBytes = imageObject.GetImageBytes();

            using var imageStream = new MemoryStream(imageBytes);
            using var image = System.Drawing.Image.FromStream(imageStream);
            string fileName = $"image_{Guid.NewGuid()}.png";
            string filePath = Path.Combine(outputDir, fileName);
            image.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
        }
    }

    public ICollection<EventType> GetSupportedEvents()
    {
        return [EventType.RENDER_IMAGE];
    }
}