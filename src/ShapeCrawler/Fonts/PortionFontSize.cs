using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services.Factories;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shared;
using ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Fonts;

internal class PortionFontSize : IFontSize
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Text aText;

    internal PortionFontSize(TypedOpenXmlPart sdkTypedOpenXmlPart, A.Text aText)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aText = aText;
    }
    
    int IFontSize.Size()
    {
        var fontSize = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize
            ?.Value;
        if (fontSize != null)
        {
            return fontSize.Value / 100;
        }

        var size = new ReferencedIndent(this.sdkTypedOpenXmlPart, this.aText).FontSizeOrNull();
        if (size != null)
        {
            return size.Value;
        }

        var indentLevel = new AParagraphWrap(this.aText.Ancestors<A.Paragraph>().First()).IndentLevel();
        var sdkSlidePart = (SlidePart)this.sdkTypedOpenXmlPart;
        
        var pPresentation = ((PresentationDocument)sdkSlidePart.OpenXmlPackage).PresentationPart!.Presentation;
        if (pPresentation.DefaultTextStyle != null)
        {
            var defaultTextStyleFonts = new IndentFonts(pPresentation.DefaultTextStyle);
            var defaultTextStyleFont = defaultTextStyleFonts.FontOrNull(indentLevel);
            if (defaultTextStyleFont.HasValue && defaultTextStyleFont.Value.Size != null)
            {
                return defaultTextStyleFont.Value.Size!.Value / 100;
            }    
        }

        var aTextDefault = pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
        if (aTextDefault is not null)
        {
            var listStyleFonts = new IndentFonts(aTextDefault.ListStyle!);
            var listStyleFont = listStyleFonts.FontOrNull(indentLevel);
            if (listStyleFont.HasValue && listStyleFont.Value.Size != null)
            {
                return listStyleFont.Value.Size!.Value / 100;
            }    
        }
        
        var indentFonts = new IndentFonts(sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!);
        var indentFont = indentFonts.FontOrNull(indentLevel);
        if (indentFont != null)
        {
            return indentFont.Value.Size!.Value / 100;
        }
        
        return Constants.DefaultFontSize;
    }

    void IFontSize.Update(int points)
    {
        var parent = this.aText.Parent!;
        var aRunPr = parent.GetFirstChild<A.RunProperties>();
        if (aRunPr == null)
        {
            var builder = new ARunPropertiesBuilder();
            aRunPr = builder.Build();
            parent.InsertAt(aRunPr, 0);
        }

        aRunPr.FontSize = points * 100;
    }
}