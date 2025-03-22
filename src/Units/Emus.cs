namespace ShapeCrawler.Units;

internal readonly ref struct Emus(long emus)
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;

    internal int AsHorizontalPixels() => (int)(emus * HorizontalResolutionDpi / 914400);
    
    internal int AsVerticalPixels() => (int)(emus * VerticalResolutionDpi / 914400);

    internal decimal AsPoints() => emus / 12700m;
}