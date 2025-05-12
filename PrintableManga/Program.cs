using System.Text.RegularExpressions;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using HtmlAgilityPack;

using DrawingPic = DocumentFormat.OpenXml.Drawing.Pictures;
using OpenDrawing = DocumentFormat.OpenXml.Drawing;
using WordProc = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace PrintableManga;

internal class Program
{
    static async Task Main(string[] args)
    {
        var outputFolder = @"E:\OnePiece Manga";

        for (int i = 1011; i <= 1148; i++)
        {
            var url = $"https://ww11.readonepiece.com/chapter/one-piece-digital-colored-comics-chapter-{i}/";

            var folderName = i.ToString().PadLeft(4, '0') + " - Colorized";
            var downloadResult = await DownloadImages(url, outputFolder, folderName, false);

            if (!downloadResult.AllSuccess)
            {
                url = $"https://ww11.readonepiece.com/chapter/one-piece-chapter-{i}/";

                downloadResult = await DownloadImages(url, outputFolder, folderName, true);
            }

            if (downloadResult.AllSuccess)
            {
                //var templateDest = Path.Combine(downloadResult.FinalOutputFolder, "Result.docx");

                //if (File.Exists(templateDest))
                //{
                //    File.Delete(templateDest);
                //}

                //GenerateMangaDoc(downloadResult.FinalOutputFolder, templateDest, maxHeightInInches: 14);
            }
        }

        //await DownloadImages(
        //    $"https://ww11.readonepiece.com/chapter/one-piece-digital-colored-comics-chapter-1026/", 
        //    outputFolder, true);
    }

    static string PrepareDirectory(string url, string outputFolder, string folderName, bool deleteIfExists)
    {
        var finalOutputFolder = Path.Combine(outputFolder, folderName);

        if (Directory.Exists(finalOutputFolder))
        {
            if (deleteIfExists)
            {
                Directory.Delete(finalOutputFolder, true);
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

    static async Task<DownloadImagesResult> DownloadImages(string url, string outputFolder, string folderName, bool deleteIfExists)
    {
        var finalOutputFolder = PrepareDirectory(url, outputFolder, folderName, deleteIfExists);

        var res = new DownloadImagesResult
        {
            AllSuccess = true,
            FinalOutputFolder = finalOutputFolder,
            ImagesCount = Directory.GetFiles(finalOutputFolder).Length,
        };

        if (res.ImagesCount > 1)
        {
            res.AllSuccess = true;

            return res;
        }

        var web = new HtmlWeb();
        var doc = await web.LoadFromWebAsync(url);

        // Select the main container
        var container = doc.DocumentNode.SelectSingleNode("//div[contains(@class, 'js-pages-container')]");

        if (container == null)
        {
            res.AllSuccess = false;
            Console.WriteLine("Main container not found.");
            return res;
        }

        // Select all img elements within the container
        var imgNodes = container.SelectNodes(".//div/img")?
            .Select(a => a.GetAttributeValue("src", string.Empty).Trim())
            .Where(a => !string.IsNullOrEmpty(a))
            .ToArray() ?? Array.Empty<string>();

        if (!imgNodes.Any())
        {
            res.AllSuccess = false;
            Console.WriteLine("No images found.");
            return res;
        }       

        using var httpClient = new HttpClient();

        for (int i = 0; i < imgNodes.Length; i++)
        {
            var imgSrc = imgNodes[i];

            try
            {
                Console.WriteLine($"Downloading: {imgSrc}");

                // Download the image
                var imageBytes = await httpClient.GetByteArrayAsync(imgSrc);

                var imgModel = ImageModel.FromBytes(imageBytes, imgSrc);

                // Save the image to the output folder
                var fileName = GetFileName(imgModel, i);
                var filePath = Path.Combine(res.FinalOutputFolder, fileName);

                await File.WriteAllBytesAsync(filePath, imageBytes);

                Console.WriteLine($"Saved: {filePath}");

                res.ImagesCount++;

                res.AllSuccess &= true;
            }
            catch (Exception ex)
            {
                res.AllSuccess &= false;
                Console.WriteLine($"Failed to download {imgSrc}: {ex.Message}");
            }
        }

        Console.WriteLine("Download complete.");

        return res;
    }

    static string GetFileName(ImageModel imageModel, int currentImgIx)
    {
        return $"{currentImgIx + 1}.{imageModel.FileExtension}";
    }

    static WordprocessingDocument CreateDocument(string templatePath)
    {
        var document = WordprocessingDocument.Create(templatePath, WordprocessingDocumentType.Document);

        var mainPart = document.AddMainDocumentPart();
        mainPart.Document = new Document(new Body(new SectionProperties(
            new PageSize
            {
                Width = 16838, // A4 width in twentieths of a point (11906 for portrait, 16838 for landscape)
                Height = 11906, // A4 height in twentieths of a point (16838 for portrait, 11906 for landscape)
                Orient = PageOrientationValues.Landscape
            },
            new PageMargin
            {
                Top = 360, // 0.25 inch
                Bottom = 360, // 0.25 inch
                Left = 360, // 0.25 inch
                Right = 360 // 0.25 inch
            }
        )));

        return document;
    }

    static void GenerateMangaDoc(string finalOutputFolder, string templatePath, double? maxHeightInInches = null)
    {
        var imageFiles = Directory.GetFiles(finalOutputFolder, "*.*")
                            .OrderBy(a => int.Parse(Path.GetFileNameWithoutExtension(a)))
                            .ToArray();

        if (!imageFiles.Any())
        {
            Console.WriteLine("No images found to add to the document.");
            return;
        }

        var document = CreateDocument(templatePath);

        var mainPart = document.MainDocumentPart!;
        var elementsBody = mainPart.Document.Body!;

        var cellQueue = new Queue<TableCell>();

        // loop all images, if its horizontal append it to the body
        // if its vertical, create table with 1 row 2 columns and append it to the cell
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
        document.Dispose();
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

        using (var stream = new FileStream(imageModel.Path!, FileMode.Open))
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