using ShapeCrawler.Fonts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Wrappers;

internal sealed record SDKPSlideMasterWrap
{
    private readonly P.SlideMaster pSlideMaster;

    internal SDKPSlideMasterWrap(P.SlideMaster pSlideMaster)
    {
        this.pSlideMaster = pSlideMaster;
    }

    internal ParagraphLevelFont? BodyStyleFontOrNull(int paraLevel)
    {
        return new ParagraphLevelFonts(this.pSlideMaster.TextStyles!.BodyStyle!).FontOrNull(paraLevel);
    }
}