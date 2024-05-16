using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

internal sealed class SlideSize
{
    private readonly P.SlideSize pSlideSize;

    internal SlideSize(P.SlideSize pSlideSize)
    {
        this.pSlideSize = pSlideSize;
    }

    internal decimal Width() => UnitConverter.HorizontalEmuToPixel(this.pSlideSize.Cx!.Value);

    internal decimal Height() => UnitConverter.HorizontalEmuToPixel(this.pSlideSize.Cy!.Value);

    internal void UpdateWidth(decimal pixels)
    {
        var emu = UnitConverter.HorizontalPixelToEmu(pixels);
        this.pSlideSize.Cx = new Int32Value((int)emu);
    }
    
    internal void UpdateHeight(decimal pixels)
    {
        var emu = UnitConverter.VerticalPixelToEmu(pixels);
        this.pSlideSize.Cy = new Int32Value((int)emu);
    }

}