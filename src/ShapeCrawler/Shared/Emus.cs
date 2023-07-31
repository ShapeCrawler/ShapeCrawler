namespace ShapeCrawler.Shared;

internal readonly ref struct Emus
{
    private readonly long emu;
    private const int HorizontalResolutionDpi = 96;
    private const int VerticalResolutionDpi = 96;

    internal Emus(long emus)
    {
        this.emu = emus;
    }

    internal int AsHorizontalPixels()
    {
        return (int)(this.emu * HorizontalResolutionDpi / 914400);
    }
    
    internal int AsVerticalPixels()
    {
        return (int)(this.emu * VerticalResolutionDpi / 914400);
    }
}