using ShapeCrawler.Fonts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

internal readonly ref struct WrappedPSlideMaster
{
    private readonly P.SlideMaster pSlideMaster;

    internal WrappedPSlideMaster(P.SlideMaster pSlideMaster)
    {
        this.pSlideMaster = pSlideMaster;
    }

    internal IndentFont? BodyStyleFontOrNull(int paraLevel) =>
        new IndentFonts(this.pSlideMaster.TextStyles!.BodyStyle!).FontOrNull(paraLevel);
}