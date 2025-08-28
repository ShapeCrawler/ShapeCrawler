namespace Fixture;

public sealed class ImageOptions
{
    public string? FormatName { get; private set; }

    public void Format(string format)
    {
        this.FormatName = format;
    }
}