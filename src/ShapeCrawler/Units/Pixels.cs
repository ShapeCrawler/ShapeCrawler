namespace ShapeCrawler.Units;

internal readonly ref struct Pixels(decimal pixels)
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private const int EmusPerInch = 914400;

    internal long AsHorizontalEmus() => (long)(pixels * EmusPerInch / HorizontalResolutionDpi);
   
    internal long AsVerticalEmus() => (long)(pixels * EmusPerInch / VerticalResolutionDpi);

    internal decimal AsPoints() => pixels * 72 / 96;
}