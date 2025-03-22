using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Fonts;

internal class PortionFontSize(A.Text aText): IFontSize
{
    decimal IFontSize.Size
    {
        get
        {
            // Try getting font size from run properties
            var runPropertiesFontSize = this.GetFontSizeFromRunProperties();
            if (runPropertiesFontSize.HasValue)
            {
                return runPropertiesFontSize.Value;
            }

            // Try getting font size from referenced indent level
            var referencedIndentFontSize = this.GetFontSizeFromReferencedIndentLevel();
            if (referencedIndentFontSize.HasValue)
            {
                return referencedIndentFontSize.Value;
            }

            // Try getting font size from slide master and shapes
            var indentLevel = new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel();
            var slideMasterPart = this.GetSlideMasterPart();
            
            // Try placeholder shapes
            var parentAutoShape = this.GetParentAutoShape();
            if (parentAutoShape != null)
            {
                var placeholderFontSize = this.GetFontSizeFromPlaceholder(parentAutoShape, slideMasterPart, indentLevel);
                if (placeholderFontSize.HasValue)
                {
                    return placeholderFontSize.Value;
                }
            }

            // Try presentation default styles
            var presentationFontSize = this.GetFontSizeFromPresentationDefaults(indentLevel);
            if (presentationFontSize.HasValue)
            {
                return presentationFontSize.Value;
            }

            // Try slide master body style (second attempt)
            if (parentAutoShape?.IsPlaceholder == true)
            {
                var bodyStyleFontSize = this.GetFontSizeFromBodyStyle(slideMasterPart, indentLevel);
                if (bodyStyleFontSize.HasValue)
                {
                    return bodyStyleFontSize.Value;
                }
            }

            // Try theme object defaults
            var themeDefaultFontSize = this.GetFontSizeFromThemeDefaults(indentLevel);
            if (themeDefaultFontSize.HasValue)
            {
                return themeDefaultFontSize.Value;
            }

            return 18; // default: https://bit.ly/37Tjjlo
        }

        set
        {
            var parent = aText.Parent!;
            var aRunPr = parent.GetFirstChild<A.RunProperties>() ?? 
                parent.InsertAt(new A.RunProperties { Language = "en-US", Dirty = false }, 0);
            aRunPr.FontSize = (int)(value * 100);
        }
    }

    private decimal? GetFontSizeFromRunProperties()
    {
        var hundredsPoints = aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize?.Value;
        return hundredsPoints.HasValue ? hundredsPoints.Value / 100m : null;
    }

    private decimal? GetFontSizeFromReferencedIndentLevel()
    {
        var hundredsPoints = new ReferencedIndentLevel(aText).FontSizeOrNull();
        return hundredsPoints.HasValue ? hundredsPoints.Value / 100m : null;
    }

    private SlideMasterPart GetSlideMasterPart()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        return openXmlPart is SlideMasterPart slideMasterPart
            ? slideMasterPart
            : ((SlidePart)openXmlPart).SlideLayoutPart!.SlideMasterPart!;
    }

    private AutoShape? GetParentAutoShape()
    {
        var parentPShape = aText.Ancestors<P.Shape>().FirstOrDefault();
        if (parentPShape is null)
        {
            return null;
        }
        
        return new AutoShape(aText.Ancestors<P.Shape>().First());
    }

    private decimal? GetFontSizeFromPlaceholder(AutoShape autoShape, SlideMasterPart slideMasterPart, int indentLevel)
    {
        if (!autoShape.IsPlaceholder)
        {
            return null;
        }

        // Check if it's a title placeholder
        if (autoShape.PlaceholderType == PlaceholderType.Title)
        {
            var titleFontSizeHundredPoints = slideMasterPart.SlideMaster.TextStyles!
                .TitleStyle!.Level1ParagraphProperties!
                .GetFirstChild<A.DefaultRunProperties>() !.FontSize!.Value;

            return titleFontSizeHundredPoints / 100m;
        }

        // Check body style
        var indentFonts = new IndentFonts(slideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
        var indentFont = indentFonts.FontOrNull(indentLevel);
        if (indentFont?.Size != null)
        {
            return indentFont.Value.Size!.Value / 100m;
        }

        return null;
    }

    private decimal? GetFontSizeFromPresentationDefaults(int indentLevel)
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var pPresentation = ((PresentationDocument)openXmlPart.OpenXmlPackage).PresentationPart!.Presentation;
        
        if (pPresentation.DefaultTextStyle == null)
        {
            return null;
        }

        // Try default text style
        var defaultTextStyleFonts = new IndentFonts(pPresentation.DefaultTextStyle);
        var defaultTextStyleFont = defaultTextStyleFonts.FontOrNull(indentLevel);
        if (defaultTextStyleFont?.Size != null)
        {
            return defaultTextStyleFont.Value.Size!.Value / 100m;
        }

        // Try theme part text defaults
        var aTextDefault = pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
        if (aTextDefault?.ListStyle != null)
        {
            var listStyleFonts = new IndentFonts(aTextDefault.ListStyle);
            var listStyleFontsFont = listStyleFonts.FontOrNull(indentLevel);
            if (listStyleFontsFont?.Size != null)
            {
                return listStyleFontsFont.Value.Size!.Value / 100m;
            }
        }

        return null;
    }

    private decimal? GetFontSizeFromBodyStyle(SlideMasterPart slideMasterPart, int indentLevel)
    {
        var indentFonts = new IndentFonts(slideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
        var indentFont = indentFonts.FontOrNull(indentLevel);
        return indentFont?.Size != null ? indentFont.Value.Size!.Value / 100m : null;
    }

    private decimal? GetFontSizeFromThemeDefaults(int indentLevel)
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var pPresentation = ((PresentationDocument)openXmlPart.OpenXmlPackage).PresentationPart!;
        var aTextDefault = pPresentation.ThemePart!.Theme.ObjectDefaults!.TextDefault;
        
        if (aTextDefault?.ListStyle == null)
        {
            return null;
        }
        
        var listStyleFonts = new IndentFonts(aTextDefault.ListStyle);
        var listStyleFont = listStyleFonts.FontOrNull(indentLevel);
        return listStyleFont?.Size != null ? listStyleFont.Value.Size!.Value / 100m : null;
    }
}