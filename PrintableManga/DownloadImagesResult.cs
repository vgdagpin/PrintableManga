namespace PrintableManga;

public class DownloadImagesResult
{
    public int ImagesCount { get; set; }

    public string FinalOutputFolder { get; set; } = null!;

    public bool AllSuccess { get; set; }
}
