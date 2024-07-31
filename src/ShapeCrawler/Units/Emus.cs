namespace ShapeCrawler.Units;

internal readonly ref struct Emus
{
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;
    private const int EmusPerInch = 914400;
    private const float EmusPerPoint = 12700;
    private readonly long value;

    internal Emus(long value)
    {
        this.value = value;
    }

    internal int AsHorizontalPixels() => (int)(this.value * HorizontalResolutionDpi / EmusPerInch);
    
    internal int AsVerticalPixels() => (int)(this.value * VerticalResolutionDpi / EmusPerInch);

    internal float AsPoints() => this.value / EmusPerPoint;
}