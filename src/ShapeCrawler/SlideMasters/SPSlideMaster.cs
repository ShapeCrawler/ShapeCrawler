using ShapeCrawler.Fonts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

// ReSharper disable once InconsistentNaming
internal readonly ref struct SPSlideMaster
{
    private readonly P.SlideMaster pSlideMaster;

    internal SPSlideMaster(P.SlideMaster pSlideMaster)
    {
        this.pSlideMaster = pSlideMaster;
    }

    internal IndentFont? BodyStyleFontOrNull(int paraLevel) =>
        new IndentFonts(this.pSlideMaster.TextStyles!.BodyStyle!).FontOrNull(paraLevel);
}