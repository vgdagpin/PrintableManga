using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using HtmlAgilityPack;

using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

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
            var templateDest = Path.Combine(finalOutputFolder, "Result.docx");

            File.Copy(mangaTemplatePath, templateDest, true);

            GenerateMangaDoc(templateDest);
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


    static void GenerateMangaDoc(string templatePath)
    {
        using (var document = WordprocessingDocument.Open(templatePath, true))
        {
            var elementsBody = document.MainDocumentPart!.Document.Body;

            var table = elementsBody!.Elements<Table>().ElementAt(0);
            var tableRow = table.Elements<TableRow>().ElementAt(0);
            
            var mainPart = document.MainDocumentPart;

            AddImageToCell(mainPart, tableRow.Elements<TableCell>().ElementAt(0), @"E:\OnePiece Manga\one-piece-chapter-1148\Onepiece_1148_t_001.png");
            AddImageToCell(mainPart, tableRow.Elements<TableCell>().ElementAt(1), @"E:\OnePiece Manga\one-piece-chapter-1148\Onepiece_1148_t_002.png");

            document.Save();
        }
    }

    private static void AddImageToCell(MainDocumentPart mainPart, TableCell cell, string imagePath)
    {
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);

        using (var stream = new FileStream(imagePath, FileMode.Open))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = mainPart.GetIdOfPart(imagePart);
        var imageModel = ImageModel.FromPath(imagePath);

        var element = new Drawing(
          new DW.Inline(
              new DW.Extent()
              {
                  Cx = imageModel.WidthInEmus / 2,
                  Cy = imageModel.HeightInEmus / 2
              },
              new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
              new DW.DocProperties() { Id = 2U, Name = "Picture 1" },
              new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = true, NoResize = false, NoSelection = false }),
              new A.Graphic(new A.GraphicData(new PIC.Picture
                        (
                          new PIC.NonVisualPictureProperties(
                              new PIC.NonVisualDrawingProperties()
                              {
                                  Id = 1U,
                                  Name = "image.png"
                              },
                              new PIC.NonVisualPictureDrawingProperties()),
                          new PIC.BlipFill(
                              new A.Blip(
                                  new A.BlipExtensionList(
                                      new A.BlipExtension()
                                      {
                                          Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                      }))
                              {
                                  Embed = relationshipId,
                                  CompressionState = A.BlipCompressionValues.Print
                              },
                              new A.Stretch(new A.FillRectangle())),
                          new PIC.ShapeProperties(
                              new A.Transform2D(new A.Offset() { X = 0L, Y = 0L }, new A.Extents() { Cx = imageModel.WidthInEmus, Cy = imageModel.HeightInEmus }),
                              new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })
                          ))
              {
                  Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
              })));

        cell.Append(new Paragraph(new Run(element)));
    }
}