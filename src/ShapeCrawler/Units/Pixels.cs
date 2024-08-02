namespace ShapeCrawler.Units;

internal readonly ref struct Pixels
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private const int EmusPerInch = 914400;
    private readonly decimal pixels;

    internal Pixels(decimal pixels)
    {
        this.pixels = pixels;
    }

    internal long AsHorizontalEmus() => (long)(this.pixels * EmusPerInch / HorizontalResolutionDpi);
   
    internal long AsVerticalEmus() => (long)(this.pixels * EmusPerInch / VerticalResolutionDpi);
}