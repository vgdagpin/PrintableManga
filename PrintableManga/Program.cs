using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using HtmlAgilityPack;

using OpenDrawing = DocumentFormat.OpenXml.Drawing;
using WordProc = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DrawingPic = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Drawing;
using System.Text.RegularExpressions;

namespace PrintableManga;

internal class Program
{
    static async Task Main(string[] args)
    {
        var url = "https://ww11.readonepiece.com/chapter/one-piece-chapter-1000/";
        var outputFolder = @"C:\Users\vince\OneDrive\OnePiece Manga";
        var mangaTemplatePath = Path.Combine(outputFolder, "Printer Comic Template.docx");

        var finalOutputFolder = await DownloadImages(url, outputFolder, false);

        if (finalOutputFolder != null)
        {
            var templateDest = Path.Combine(finalOutputFolder, "Result.docx");

            File.Copy(mangaTemplatePath, templateDest, true);

            GenerateMangaDoc(finalOutputFolder, templateDest);
        }
    }

    static string PrepareDirectory(string url, string outputFolder, bool deleteIfExists)
    {
        var chapter = Regex.Match(url.Trim('/').Split('/').Last(), @"\d+$").Value.PadLeft(4, '0');

        var finalOutputFolder = Path.Combine(outputFolder, chapter);

        if (Directory.Exists(finalOutputFolder))
        {
            if (deleteIfExists)
            {
                Directory.Delete(finalOutputFolder);
            }
            else
            {
                return finalOutputFolder;
            }
        }

        // Ensure the output folder exists
        Directory.CreateDirectory(finalOutputFolder);

        return finalOutputFolder;
    }

    static async Task<string?> DownloadImages(string url, string outputFolder, bool deleteIfExists)
    {
        var finalOutputFolder = PrepareDirectory(url, outputFolder, deleteIfExists);

        if (Directory.GetFiles(finalOutputFolder).Length > 1)
        {
            return finalOutputFolder;
        }

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

            if (!string.IsNullOrEmpty(imgSrc))
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
        using (var document = WordprocessingDocument.Open(templatePath, true))
        {
            var elementsBody = document.MainDocumentPart!.Document.Body;

            var table = elementsBody!.Elements<Table>().First();

            var mainPart = document.MainDocumentPart;

            var imageFiles = Directory.GetFiles(finalOutputFolder, "*.*")
                                                  .Where(file => !file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                                                  .ToArray();

            if (!imageFiles.Any())
            {
                Console.WriteLine("No images found to add to the document.");
                return;
            }

            var ix = 0;
            foreach (var cell in GetCells(table).Skip(1))
            {
                AddImageToCell(mainPart, cell, imageFiles[ix]);

                ix++;

                if (ix >= imageFiles.Length)
                {
                    break;
                }
            }
          
            document.Save();
        }
    }

    static IEnumerable<TableCell> GetCells(Table table)
    {
        foreach (var row in table.Elements<TableRow>())
        {
            foreach (var cell in row.Elements<TableCell>())
            {
                yield return cell;
            }
        }
    }

    private static void AddImageToCell(MainDocumentPart mainPart, TableCell cell, string imagePath)
    {
        var imageModel = ImageModel.FromPath(imagePath, maxHeightInInches: 14);
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);

        using (var stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = mainPart.GetIdOfPart(imagePart);

        var element = new Drawing
            (
                new WordProc.Inline
                (
                    new WordProc.Extent
                    {
                        Cx = imageModel.WidthInEmus / 2,
                        Cy = imageModel.HeightInEmus / 2
                    },
                    new WordProc.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                    new WordProc.DocProperties() { Id = 2U, Name = "Picture 1" },
                    new WordProc.NonVisualGraphicFrameDrawingProperties(new OpenDrawing.GraphicFrameLocks() { NoChangeAspect = true, NoResize = false, NoSelection = false }),
                    new OpenDrawing.Graphic(new OpenDrawing.GraphicData(new DrawingPic.Picture
                        (
                            new DrawingPic.NonVisualPictureProperties(
                                new DrawingPic.NonVisualDrawingProperties
                                {
                                    Id = 1U,
                                    Name = "image.png"
                                },
                                new DrawingPic.NonVisualPictureDrawingProperties()),
                            new DrawingPic.BlipFill(
                                new OpenDrawing.Blip
                                {
                                    Embed = relationshipId,
                                    CompressionState = OpenDrawing.BlipCompressionValues.Print
                                },
                                new OpenDrawing.Stretch(new OpenDrawing.FillRectangle())),
                            new DrawingPic.ShapeProperties(
                                new OpenDrawing.Transform2D(new OpenDrawing.Offset() { X = 0L, Y = 0L }, new OpenDrawing.Extents() { Cx = imageModel.WidthInEmus, Cy = imageModel.HeightInEmus }),
                                new OpenDrawing.PresetGeometry(new OpenDrawing.AdjustValueList()) { Preset = OpenDrawing.ShapeTypeValues.Rectangle })
                        ))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    })
                )
            );

        var paragraph = new Paragraph(
            new ParagraphProperties(new Justification()
            {
                Val = JustificationValues.Center  // Right-align the content
            }),
            new Run(element));

        cell.Append(paragraph);
    }
}

public class ImageModel
{
    public long WidthInEmus { get; set; }
    public long HeightInEmus { get; set; }

    public static ImageModel FromPath(string path, double? maxHeightInInches = null)
    {
        var result = new ImageModel();

#pragma warning disable CA1416 // Validate platform compatibility
        using (var img = Image.FromFile(path))
        {
            // Calculate EMUs considering image DPI
            var hResolution = img.HorizontalResolution > 0 ? img.HorizontalResolution : 96;
            var vResolution = img.VerticalResolution > 0 ? img.VerticalResolution : 96;

            double originalWidthInInches = img.Width / hResolution;
            double originalHeightInInches = img.Height / vResolution;

            if (maxHeightInInches.HasValue)
            {
                // Scale dimensions to fit within maxHeightInInches while maintaining aspect ratio
                double scaleFactor = maxHeightInInches.Value / originalHeightInInches;

                double adjustedHeightInInches = maxHeightInInches.Value;
                double adjustedWidthInInches = originalWidthInInches * scaleFactor;

                result.WidthInEmus = (long)(adjustedWidthInInches * 914400);
                result.HeightInEmus = (long)(adjustedHeightInInches * 914400);
            }
            else
            {
                // Use original dimensions if no maxHeightInInches is provided
                result.WidthInEmus = (long)(originalWidthInInches * 914400);
                result.HeightInEmus = (long)(originalHeightInInches * 914400);
            }
        }
#pragma warning restore CA1416 // Validate platform compatibility

        return result;
    }
}
