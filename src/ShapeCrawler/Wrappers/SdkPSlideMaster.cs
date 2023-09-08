using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed record SdkPSlideMaster
{
    private readonly P.SlideMaster pSlideMaster;

    internal SdkPSlideMaster(P.SlideMaster pSlideMaster)
    {
        this.pSlideMaster = pSlideMaster;
    }

    internal IndentFont? BodyStyleFontOrNull(int paraLevel)
    {
        return new IndentFonts(this.pSlideMaster.TextStyles!.BodyStyle!).FontOrNull(paraLevel);
    }
}