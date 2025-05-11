using System.Drawing;

namespace PrintableManga;

public class ImageModel
{
    public long WidthInEmus { get; set; }
    public long HeightInEmus { get; set; }

    public static ImageModel FromPath(string path)
    {
        var result = new ImageModel();

        using (var img = Image.FromFile(path))
        {
            // Calculate EMUs considering image DPI
            var hResolution = img.HorizontalResolution > 0 ? img.HorizontalResolution : 96;
            var vResolution = img.VerticalResolution > 0 ? img.VerticalResolution : 96;

            double widthInInches = img.Width / hResolution;
            double heightInInches = img.Height / vResolution;

            result.WidthInEmus = (long)(widthInInches * 914400);
            result.HeightInEmus = (long)(heightInInches * 914400);            
        }

        return result;
    }
}
