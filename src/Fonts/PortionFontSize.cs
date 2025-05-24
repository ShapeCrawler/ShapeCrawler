using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Fonts;

internal class PortionFontSize(A.Text aText) : IFontSize
{
    decimal IFontSize.Size
    {
        get
        {
            var runPropertiesSize = this.GetRunPropertiesFontSizeOrNull();
            if (runPropertiesSize.HasValue)
            {
                return this.ApplyNormAutofitScaling(runPropertiesSize.Value);
            }

            return this.GetBodyStyleFontSize()
                   ?? this.GetPlaceholderFontSize()
                   ?? this.GetPresentationDefaultsFontSize()
                   ?? this.GetThemeDefaultsFontSize()
                   ?? 18; // default: https://bit.ly/37Tjjlo
        }

        set
        {
            var parent = aText.Parent!;
            var aRunPr = parent.GetFirstChild<A.RunProperties>() ??
                         parent.InsertAt(new A.RunProperties { Language = "en-US", Dirty = false }, 0);
            aRunPr.FontSize = (int)(value * 100);
        }
    }

    private decimal ApplyNormAutofitScaling(decimal size)
    {
        var bodyPr = aText.Ancestors<A.Paragraph>().First().Ancestors<P.TextBody>().FirstOrDefault()
            ?.GetFirstChild<A.BodyProperties>();
        var normAutofit = bodyPr?.GetFirstChild<A.NormalAutoFit>();
        if (normAutofit?.FontScale != null)
        {
            decimal fontScale = normAutofit.FontScale.Value / 100000m;
            size = Math.Round(size * fontScale, MidpointRounding.AwayFromZero);
        }

        return size;
    }

    private decimal? GetBodyStyleFontSize()
    {
        var indentLevel = new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel();
        var slideMasterPart = this.GetSlideMasterPart();

        var indentFonts = new IndentFonts(slideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
        var indentFont = indentFonts.FontOrNull(indentLevel);
        return indentFont?.Size != null ? indentFont.Value.Size!.Value / 100m : null;
    }

    private decimal? GetPlaceholderFontSize()
    {
        var parentShape = this.GetParentShapeOrNull();
        if (parentShape == null)
        {
            return null;
        }

        var slideMasterPart = this.GetSlideMasterPart();

        if (parentShape.PlaceholderType == PlaceholderType.Title)
        {
            var titleFontSizeHundredPoints = slideMasterPart.SlideMaster.TextStyles!
                .TitleStyle!.Level1ParagraphProperties!
                .GetFirstChild<A.DefaultRunProperties>() !.FontSize!.Value;

            return titleFontSizeHundredPoints / 100m;
        }

        var indentFonts = new IndentFonts(slideMasterPart.SlideMaster.TextStyles!.BodyStyle!);
        var indentFont = indentFonts.FontOrNull(new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel());
        return indentFont?.Size != null ? indentFont.Value.Size!.Value / 100m : null;
    }

    private decimal? GetPresentationDefaultsFontSize()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var pPresentation = ((PresentationDocument)openXmlPart.OpenXmlPackage).PresentationPart!.Presentation;

        if (pPresentation.DefaultTextStyle == null)
        {
            return null;
        }

        var defaultTextStyleFonts = new IndentFonts(pPresentation.DefaultTextStyle);
        var defaultTextStyleFont = defaultTextStyleFonts.FontOrNull(new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel());
        if (defaultTextStyleFont?.Size != null)
        {
            return defaultTextStyleFont.Value.Size!.Value / 100m;
        }

        var aTextDefault = pPresentation.PresentationPart!.ThemePart!.Theme.ObjectDefaults!.TextDefault;
        if (aTextDefault?.ListStyle != null)
        {
            var listStyleFonts = new IndentFonts(aTextDefault.ListStyle);
            var listStyleFontsFont = listStyleFonts.FontOrNull(new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel());
            if (listStyleFontsFont?.Size != null)
            {
                return listStyleFontsFont.Value.Size!.Value / 100m;
            }
        }

        return null;
    }

    private decimal? GetThemeDefaultsFontSize()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var pPresentation = ((PresentationDocument)openXmlPart.OpenXmlPackage).PresentationPart!;
        var aTextDefault = pPresentation.ThemePart!.Theme.ObjectDefaults!.TextDefault;

        if (aTextDefault?.ListStyle == null)
        {
            return null;
        }

        var listStyleFonts = new IndentFonts(aTextDefault.ListStyle);
        var listStyleFont = listStyleFonts.FontOrNull(new SCAParagraph(aText.Ancestors<A.Paragraph>().First()).GetIndentLevel());
        return listStyleFont?.Size != null ? listStyleFont.Value.Size!.Value / 100m : null;
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

        var pShape = aText.Ancestors<P.Shape>().First();

        return new Shape(new Position(pShape), new ShapeSize(pShape), new ShapeId(pShape), pShape);
    }
}