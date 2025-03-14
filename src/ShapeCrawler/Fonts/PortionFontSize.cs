using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Fonts;

internal class PortionFontSize : IFontSize
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Text aText;

    internal PortionFontSize(OpenXmlPart sdkTypedOpenXmlPart, A.Text aText)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aText = aText;
    }

    float IFontSize.Size()
    {
        var hundredsOfPoint = this.aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize
            ?.Value;
        if (hundredsOfPoint != null)
        {
            return hundredsOfPoint.Value / 100f;
        }
 
        hundredsOfPoint = new ReferencedIndentLevel(this.sdkTypedOpenXmlPart, this.aText).FontSizeOrNull();
        if (hundredsOfPoint is not null)
        {
            return hundredsOfPoint.Value / 100f; 
        }

        var indentLevel = new SCAParagraph(this.aText.Ancestors<A.Paragraph>().First()).IndentLevel();
        SlideMasterPart sdkSlideMasterPart;
        if (this.sdkTypedOpenXmlPart is SlideMasterPart)
        {
            sdkSlideMasterPart = (SlideMasterPart)this.sdkTypedOpenXmlPart;
        }
        else
        {
            sdkSlideMasterPart = ((SlidePart)this.sdkTypedOpenXmlPart).SlideLayoutPart!.SlideMasterPart!;
        }

        AutoShape? parentAutoShape = null;
        var parentPShape = this.aText.Ancestors<P.Shape>().FirstOrDefault();
        if(parentPShape is not null)
        {
            parentAutoShape = new AutoShape(this.sdkTypedOpenXmlPart, this.aText.Ancestors<P.Shape>().First());
        }
        
        if (parentAutoShape is not null && parentAutoShape.IsPlaceholder)
        {
            if (parentAutoShape.PlaceholderType == PlaceholderType.Title)
            {
                var titleFontSize = sdkSlideMasterPart.SlideMaster.TextStyles!.TitleStyle!.Level1ParagraphProperties!
                    .GetFirstChild<A.DefaultRunProperties>() !.FontSize!.Value;
                
                return titleFontSize / 100f;
            }
            
            var indentFonts =
                new IndentFonts(sdkSlideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
            var indentFont = indentFonts.FontOrNull(indentLevel);
            if (indentFont != null)
            {
                return indentFont.Value.Size!.Value / 100f;
            }    
        }
        
        // Presentation
        var pPresentation = ((PresentationDocument)this.sdkTypedOpenXmlPart.OpenXmlPackage).PresentationPart!.Presentation;
        if (pPresentation.DefaultTextStyle != null)
        {
            var defaultTextStyleFonts = new IndentFonts(pPresentation.DefaultTextStyle);
            var defaultTextStyleFont = defaultTextStyleFonts.FontOrNull(indentLevel);
            if (defaultTextStyleFont.HasValue && defaultTextStyleFont.Value.Size != null)
            {
                return defaultTextStyleFont.Value.Size!.Value / 100f;
            }

            var aTextDefault2 = pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
            if (aTextDefault2 is not null)
            {
                var listStyleFonts =
                    new IndentFonts(pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault!.ListStyle!);
                var listStyleFontsFont = listStyleFonts.FontOrNull(indentLevel);
                if (listStyleFontsFont.HasValue && listStyleFontsFont.Value.Size != null)
                {
                    return listStyleFontsFont.Value.Size!.Value / 100f;
                }
            }
        }
        
        if (parentAutoShape is not null && parentAutoShape.IsPlaceholder)
        {
            var indentFonts2 =
                new IndentFonts(sdkSlideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
            var indentFont2 = indentFonts2.FontOrNull(indentLevel);
            if (indentFont2 != null && indentFont2.Value.Size != null)
            {
                return indentFont2.Value.Size!.Value / 100f;
            }
        }

        var aTextDefault = pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
        if (aTextDefault is not null)
        {
            var listStyleFonts = new IndentFonts(aTextDefault.ListStyle!);
            var listStyleFont = listStyleFonts.FontOrNull(indentLevel);
            if (listStyleFont.HasValue && listStyleFont.Value.Size != null)
            {
                return listStyleFont.Value.Size!.Value / 100f;
            }
        }
        
        return 18; // default: https://bit.ly/37Tjjlo
    }

    void IFontSize.Update(float points)
    {
        var parent = this.aText.Parent!;
        var aRunPr = parent.GetFirstChild<A.RunProperties>();
        if (aRunPr == null)
        {
            aRunPr = new A.RunProperties { Language = "en-US", FontSize = 1400, Dirty = false };
            parent.InsertAt(aRunPr, 0);
        }

        aRunPr.FontSize = (int)(points * 100);
    }
}