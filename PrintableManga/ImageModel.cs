using System.Drawing;

using MimeDetective;

namespace PrintableManga;

public class ImageModel
{
    static IContentInspector? inspector;
    static IContentInspector Inspector
    {
        get
        {
            inspector ??= new ContentInspectorBuilder()
            {
                Definitions = new MimeDetective.Definitions.CondensedBuilder()
                {
                    UsageType = MimeDetective.Definitions.Licensing.UsageType.PersonalNonCommercial
                }.Build()
            }.Build();

            return inspector;
        }
    }

    public long WidthInEmus { get; set; }
    public long HeightInEmus { get; set; }

    public double WidthInInches { get; set; }
    public double HeightInInches { get; set; }

    public bool IsHorizontal => WidthInInches > HeightInInches;

    public string? Path { get; set; }

    public string? FileExtension { get; set; }

    public string? MimeType { get; set; }

    public static ImageModel FromBytes(byte[] imageBytes, string imgSrc)
    {
        var result = new ImageModel();

        var defMatch = Inspector.Inspect(imageBytes)
            .OrderByDescending(a => a.Points)
            .FirstOrDefault();

        var fileDef = defMatch?.Definition;

        result.MimeType = fileDef?.File.MimeType;
        result.FileExtension = fileDef?.File.Extensions.FirstOrDefault()?.ToLower();
        result.Path = imgSrc;

        return result;
    }

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
