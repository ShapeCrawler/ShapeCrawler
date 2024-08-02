namespace ShapeCrawler.Units;

internal readonly ref struct Emus
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private readonly long value;

    internal Emus(long value)
    {
        this.value = value;
    }

    internal int AsHorizontalPixels() => (int)(this.value * HorizontalResolutionDpi / 914400);
    
    internal int AsVerticalPixels() => (int)(this.value * VerticalResolutionDpi / 914400);

    internal float AsPoints() => this.value / 12700f;
}