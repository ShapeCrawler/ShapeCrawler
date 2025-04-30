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
            var runPropertiesFontSize = this.GetRunPropertiesFontSizeOrNull();
            if (runPropertiesFontSize.HasValue)
            {
                return runPropertiesFontSize.Value;
            }

            // Try getting font size from referenced indent level
            var referencedIndentFontSize = new ReferencedIndentLevel(aText).ReferencedFontSizeOrNull();
            if (referencedIndentFontSize.HasValue)
            {
                return referencedIndentFontSize.Value;
            }

            // Try getting font size from slide master and shapes
            var indentLevel = new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel();
            var slideMasterPart = this.GetSlideMasterPart();
            
            // Try placeholder shapes
            var parentShape = this.GetParentShapeOrNull();
            if (parentShape != null)
            {
                var placeholderFontSize = GetFontSizeFromPlaceholder(parentShape, slideMasterPart, indentLevel);
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

            if (parentShape?.PlaceholderType != null)
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

    private static decimal? GetFontSizeFromPlaceholder(Shape shape, SlideMasterPart slideMasterPart, int indentLevel)
    {
        if (shape.PlaceholderType == null)
        {
            return null;   
        }

        // Check if it's a title placeholder
        if (shape.PlaceholderType == PlaceholderType.Title)
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
    
    private decimal? GetRunPropertiesFontSizeOrNull()
    {
        var hundredsPoints = aText.Parent!.GetFirstChild<A.RunProperties>()?.FontSize?.Value;
        
        return hundredsPoints / 100m;
    }

    private SlideMasterPart GetSlideMasterPart()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        
        return openXmlPart as SlideMasterPart ?? ((SlidePart)openXmlPart).SlideLayoutPart!.SlideMasterPart!;
    }

    private Shape? GetParentShapeOrNull()
    {
        var parentPShape = aText.Ancestors<P.Shape>().FirstOrDefault();
        if (parentPShape is null)
        {
            return null;
        }
        
        return new Shape(aText.Ancestors<P.Shape>().First());
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