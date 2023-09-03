using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed record SdkPSlideMasterWrap
{
    private readonly P.SlideMaster pSlideMaster;

    internal SdkPSlideMasterWrap(P.SlideMaster pSlideMaster)
    {
        this.pSlideMaster = pSlideMaster;
    }

    internal IndentFont? BodyStyleFontOrNull(int paraLevel)
    {
        return new IndentFonts(this.pSlideMaster.TextStyles!.BodyStyle!).FontOrNull(paraLevel);
    }
}