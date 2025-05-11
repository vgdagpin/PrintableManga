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
        var outputFolder = @"E:\OnePiece Manga";
        var mangaTemplatePath = Path.Combine(outputFolder, "Printer Comic Template - Blank.docx");

        var finalOutputFolder = await DownloadImages(url, outputFolder, false);

        if (finalOutputFolder != null)
        {
            var templateDest = Path.Combine(finalOutputFolder, "Result.docx");

            File.Copy(mangaTemplatePath, templateDest, true);

            GenerateMangaDoc(finalOutputFolder, templateDest, maxHeightInInches: 14);
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


    static void GenerateMangaDoc(string finalOutputFolder, string templatePath, double? maxHeightInInches = null)
    {
        using (var document = WordprocessingDocument.Open(templatePath, true))
        {
            var mainPart = document.MainDocumentPart!;
            var elementsBody = mainPart.Document.Body!;

            var imageFiles = Directory.GetFiles(finalOutputFolder, "*.*")
                                                  .Where(file => !file.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                                                  .OrderBy(a => int.Parse(Path.GetFileNameWithoutExtension(a)))
                                                  .ToArray();

            if (!imageFiles.Any())
            {
                Console.WriteLine("No images found to add to the document.");
                return;
            }

            var cellQueue = new Queue<TableCell>();

            for (int i = 0; i < imageFiles.Length; i++)
            {
                var image = imageFiles[i];

                var imageModel = ImageModel.FromPath(image, maxHeightInInches);

                var imageElement = ImageToElement(mainPart, imageModel);

                if (imageModel.IsHorizontal)
                {
                    elementsBody.Append(imageElement);

                    cellQueue = null;
                }
                else
                {
                    TableCell? cell;

                    if (cellQueue == null)
                    {
                        cellQueue = new Queue<TableCell>();

                        foreach (var eachCell in GetCells(CreateTable(elementsBody)))
                        {
                            cellQueue.Enqueue(eachCell);
                        }
                    }

                    // if still has cell, use it
                    // if not, create new table
                    if (!cellQueue.TryDequeue(out cell))
                    {
                        cellQueue = new Queue<TableCell>();

                        foreach (var eachCell in GetCells(CreateTable(elementsBody)))
                        {
                            cellQueue.Enqueue(eachCell);
                        }

                        cell = cellQueue.Dequeue();
                    }

                    cell.Append(imageElement);

                    // if this is the last image, and there's still a free cell
                    // lets just fill it with the last image
                    if (i == imageFiles.Length - 1 && cellQueue.TryDequeue(out cell))
                    {
                        cell.Append(ImageToElement(mainPart, imageModel));
                    }
                }
            }            

            document.Save();
        }
    }

    static Table CreateTable(Body body)
    {
        // Create a new table
        var table = new Table();

        // Define table properties  
        var tableProperties = new TableProperties(
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }, // Set table width to 100%  
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
            )
        );

        table.AppendChild(tableProperties);

        // Create a single row
        var tableRow = new TableRow();

        // Create two cells for the row
        var cell1 = new TableCell(new Paragraph());
        var cell2 = new TableCell(new Paragraph());

        // Add cells to the row
        tableRow.Append(cell1, cell2);

        // Add the row to the table
        table.AppendChild(tableRow);

        body.Append(table);

        return table;
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

    private static Paragraph ImageToElement(MainDocumentPart mainPart, ImageModel imageModel)
    {
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);

        using (var stream = new FileStream(imageModel.Path, FileMode.Open))
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

        return new Paragraph(
            new ParagraphProperties(new Justification()
            {
                Val = JustificationValues.Center  // Right-align the content
            }),
            new Run(element));
    }
}

public class ImageModel
{
    public long WidthInEmus { get; set; }
    public long HeightInEmus { get; set; }

    public double WidthInInches { get; set; }
    public double HeightInInches { get; set; }

    public bool IsHorizontal => WidthInInches > HeightInInches;

    public string Path { get; set; } = null!;

    public static ImageModel FromPath(string path, double? maxHeightInInches = null)
    {
        var result = new ImageModel
        {
            Path = path
        };

#pragma warning disable CA1416 // Validate platform compatibility
        using (var img = Image.FromFile(path))
        {
            // Calculate EMUs considering image DPI
            var hResolution = img.HorizontalResolution > 0 ? img.HorizontalResolution : 96;
            var vResolution = img.VerticalResolution > 0 ? img.VerticalResolution : 96;

            result.WidthInInches = img.Width / hResolution;
            result.HeightInInches = img.Height / vResolution;

            if (maxHeightInInches.HasValue)
            {
                // Scale dimensions to fit within maxHeightInInches while maintaining aspect ratio
                double scaleFactor = maxHeightInInches.Value / result.HeightInInches;

                result.HeightInInches = maxHeightInInches.Value;
                result.WidthInInches = result.WidthInInches * scaleFactor;
            }

            // Use original dimensions if no maxHeightInInches is provided
            result.WidthInEmus = (long)(result.WidthInInches * 914400);
            result.HeightInEmus = (long)(result.HeightInInches * 914400);
        }
#pragma warning restore CA1416 // Validate platform compatibility

        return result;
    }
}
