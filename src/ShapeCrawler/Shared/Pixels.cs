namespace ShapeCrawler.Shared;

internal readonly ref struct Pixels
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private readonly int pixels;

    internal Pixels(int pixels)
    {
        this.pixels = pixels;
    }

    internal long AsHorizontalEmus() => this.pixels * 914400 / HorizontalResolutionDpi;
    internal long AsVerticalEmus() => this.pixels * 914400 / VerticalResolutionDpi;
}