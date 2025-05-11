using HtmlAgilityPack;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Drawing;

namespace PrintableManga;

internal class Program
{
    static async Task Main(string[] args)
    {
        var url = "https://ww11.readonepiece.com/chapter/one-piece-chapter-1148/";
        var outputFolder = @"E:\OnePiece Manga";
        var mangaTemplatePath = Path.Combine(outputFolder, "Printer Comic Template.docx");

        var finalOutputFolder = await DownloadImages(url, outputFolder, false);

        if (finalOutputFolder != null)
        {
            var templateDest = Path.Combine(finalOutputFolder, "Printer Comic Template.docx");

            File.Copy(mangaTemplatePath, templateDest, true);

            GenerateMangaDoc(finalOutputFolder, templateDest);
        }
    }

    static async Task<string?> DownloadImages(string url, string outputFolder, bool deleteIfExists)
    {
        var chapter = url.Trim('/').Split('/').Last();

        var finalOutputFolder = Path.Combine(outputFolder, chapter);

        if (Directory.Exists(finalOutputFolder))
        {
            if (deleteIfExists)
            {
                Directory.Delete(finalOutputFolder);
            }
            else
            {
                // exit, means we already generated it
                return finalOutputFolder;
            }
        }

        // Ensure the output folder exists
        Directory.CreateDirectory(finalOutputFolder);

        var web = new HtmlWeb();
        var doc = await web.LoadFromWebAsync(url);

        // Select the main container
        var container = doc.DocumentNode.SelectSingleNode("//div[contains(@class, 'js-pages-container')]");

        if (container == null)
        {
            Console.WriteLine("Main container not found.");
            return null;
        }

        // Select all img elements within the container
        var imgNodes = container.SelectNodes(".//div/img");

        if (imgNodes == null || !imgNodes.Any())
        {
            Console.WriteLine("No images found.");
            return null;
        }

        using var httpClient = new HttpClient();

        foreach (var imgNode in imgNodes)
        {
            var imgSrc = imgNode.GetAttributeValue("src", string.Empty).Trim();

            if (imgSrc != null && imgSrc.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    Console.WriteLine($"Downloading: {imgSrc}");

                    // Download the image
                    var imageBytes = await httpClient.GetByteArrayAsync(imgSrc);

                    // Save the image to the output folder
                    var fileName = Path.GetFileName(imgSrc);
                    var filePath = Path.Combine(finalOutputFolder, fileName);

                    await File.WriteAllBytesAsync(filePath, imageBytes);

                    Console.WriteLine($"Saved: {filePath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to download {imgSrc}: {ex.Message}");
                }
            }
        }

        Console.WriteLine("Download complete.");

        return finalOutputFolder;
    }


    static void GenerateMangaDoc(string finalOutputFolder, string templatePath)
    {
        // Get all PNG files in the final output folder
        var imageFiles = Directory.GetFiles(finalOutputFolder, "*.png");

        if (!imageFiles.Any())
        {
            Console.WriteLine("No images found to add to the document.");
            return;
        }

        // Open the DOCX template
        using (var wordDoc = WordprocessingDocument.Open(templatePath, true))
        {
            var mainPart = wordDoc.MainDocumentPart;
            if (mainPart == null)
            {
                Console.WriteLine("MainDocumentPart not found in the template.");
                return;
            }

            var table = mainPart.Document.Body.Elements<Table>().FirstOrDefault();
            if (table == null)
            {
                Console.WriteLine("No table found in the document.");
                return;
            }

            // Add images to the table cells
            int imageIndex = 0;
            foreach (var row in table.Elements<TableRow>())
            {
                var cells = row.Elements<TableCell>().ToList();

                for (int i = 0; i < cells.Count && imageIndex < imageFiles.Length; i++)
                {
                    var cell = cells[i];
                    var imagePath = imageFiles[imageIndex];

                    AddImageToCell(mainPart, cell, imagePath, $"Image{imageIndex + 1}");
                    imageIndex++;
                }

                if (imageIndex >= imageFiles.Length)
                    break;
            }

            mainPart.Document.Save();
        }

        Console.WriteLine("Images added to the document successfully.");
    }

    static void AddImageToCell(MainDocumentPart mainPart, TableCell cell, string imagePath, string imageId)
    {
        // Add the image to the Word document
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using (var stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        // Get image dimensions
        using (var img = Image.FromFile(imagePath))
        {
            var widthInEmus = (long)(img.Width * 9525); // Convert pixels to EMUs
            var heightInEmus = (long)(img.Height * 9525);

            var gpxData = new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = imageId
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip
                                    {
                                        Embed = mainPart.GetIdOfPart(imagePart),
                                        CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents { Cx = widthInEmus, Cy = heightInEmus }),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList())
                                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle })));

            // Create the drawing element for the image
            var element = new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent { Cx = widthInEmus, Cy = heightInEmus },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                    {
                        Id = (UInt32Value)1U,
                        Name = imageId
                    },
                    new DocumentFormat.OpenXml.Drawing.Graphic(gpxData))
            );

            // Add the image to the cell
            var paragraph = new Paragraph(new Run(element));
            cell.Append(paragraph);
        }        
    }
}