using ShapeCrawler.Fonts;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideMasters;

// ReSharper disable once InconsistentNaming
internal readonly ref struct SCPSlideMaster
{
    private readonly P.SlideMaster pSlideMaster;

    internal SCPSlideMaster(P.SlideMaster pSlideMaster)
    {
        this.pSlideMaster = pSlideMaster;
    }

    internal IndentFont? BodyStyleFontOrNull(int paraLevel) =>
        new IndentFonts(this.pSlideMaster.TextStyles!.BodyStyle!).FontOrNull(paraLevel);
}