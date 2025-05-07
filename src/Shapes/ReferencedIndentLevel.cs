using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Colors;
using ShapeCrawler.Fonts;
using ShapeCrawler.Paragraphs;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal readonly ref struct ReferencedIndentLevel(A.Text aText)
{
    internal string? ColorHexOrNull()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        return openXmlPart switch
        {
            SlidePart => this.SlideColorHexOrNull(),
            _ => this.LayoutColorHexOrNull()
        };
    }

    internal bool? FontBoldFlagOrNull()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (openXmlPart is SlidePart)
        {
            return this.SlideFontBoldFlagOrNull();
        }

        throw new SCException("Not implemented.");
    }

    internal decimal? FontSizeOrNull()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aParagraph = aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
        var slidePShape = aText.Ancestors<P.Shape>().FirstOrDefault();
        if (slidePShape == null)
        {
            return null;
        }

        var slidePh = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return null;
        }

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNull(slidePShape);
        if (refLayoutPShapeOfSlide == null)
        {
            var refMasterPShape = this.ReferencedMasterPShapeOrNull(slidePShape);
            if (refMasterPShape != null)
            {
                var fonts = new IndentFonts(refMasterPShape!.TextBody!.ListStyle!);
                var font = fonts.FontOrNull(indentLevel);
                if (font.HasValue)
                {
                    return (int)font.Value.Size! / 100m;
                }
            }

            var sdkSlidePart = (SlidePart)openXmlPart;
            var bodyStyleFonts =
                new IndentFonts(sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!);
            var bodyStyleFont = bodyStyleFonts.FontOrNull(indentLevel);
            if (bodyStyleFont.HasValue)
            {
                return (int)bodyStyleFont.Value.Size! / 100m;
            }

            return null;
        }

        var layoutFonts = new IndentFonts(refLayoutPShapeOfSlide.TextBody!.ListStyle!);
        var layoutIndentFont = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentFont is { Size: not null })
        {
            return (int)layoutIndentFont.Value.Size! / 100m;
        }

        return this.MasterFontSizeOrNull(refLayoutPShapeOfSlide, indentLevel) / 100m;
    }

    internal A.LatinFont? ALatinFontOrNull()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        return openXmlPart switch
        {
            SlidePart slidePart => this.SlideALatinFontOrNull(slidePart),
            SlideMasterPart => this.SlideMasterALatinFont(),
            _ => throw new SCException("Not implemented.")
        };
    }

    private P.Shape? ReferencedLayoutPShapeOrNull(P.Shape pShape)
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (openXmlPart is not SlidePart slidePart)
        {
            return null;
        }

        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>() !;
        var referencedLayoutPShape =
            new SCPShapeTree(slidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!).ReferencedPShapeOrNull(
                pPlaceholderShape);

        return referencedLayoutPShape;
    }

    private P.Shape? ReferencedMasterPShapeOrNull(P.Shape pShape)
    {
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var slideOrLayoutPShapeTree = openXmlPart switch
        {
            SlidePart slidePart => slidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!,
            _ => ((SlideLayoutPart)openXmlPart).SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!
        };

        var referencedPShape = new SCPShapeTree(slideOrLayoutPShapeTree).ReferencedPShapeOrNull(pPlaceholderShape);

        return referencedPShape;
    }

    private bool HexFromName(IndentFont? indentFont, out string? referencedShapeColorOrNull)
    {
        if (indentFont == null)
        {
            referencedShapeColorOrNull = null;
            return true;
        }

        if (indentFont.Value.ARgbColorModelHex != null)
        {
            referencedShapeColorOrNull = indentFont.Value.ARgbColorModelHex.Val!.Value;
            return true;
        }

        if (indentFont.Value.ASchemeColor != null)
        {
            var presentationColor =
                new PresentationColor(aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!);
            referencedShapeColorOrNull = presentationColor.ThemeColorHex(indentFont.Value.ASchemeColor.Val!.Value);
            return true;
        }

        if (indentFont.Value.ASystemColor != null)
        {
            referencedShapeColorOrNull = indentFont.Value.ASystemColor.LastColor!;
            return true;
        }

        if (indentFont.Value.APresetColor != null)
        {
            var coloName = indentFont.Value.APresetColor.Val!.Value.ToString();
            referencedShapeColorOrNull = ColorTranslator.HexFromName(coloName);
            return true;
        }

        referencedShapeColorOrNull = null;
        return false;
    }

    private string? LayoutColorHexOrNull()
    {
        var pShape = aText.Ancestors<P.Shape>().First();
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var referencedMasterPShape = this.ReferencedMasterPShapeOrNull(pShape);
        var aParagraph = aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
        if (referencedMasterPShape != null)
        {
            var masterIndentFonts = new IndentFonts(referencedMasterPShape.TextBody!.ListStyle!);
            var masterIndentFont = masterIndentFonts.FontOrNull(indentLevel);
            if (masterIndentFont != null && this.HexFromName(masterIndentFont, out var masterColor))
            {
                return masterColor;
            }
        }

        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        return this.GetColorFromLayoutPlaceholderType(pPlaceholderShape, openXmlPart, indentLevel);
    }

    private string? GetColorFromLayoutPlaceholderType(P.PlaceholderShape pPlaceholderShape, OpenXmlPart openXmlPart, int indentLevel)
    {
        if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.Title)
        {
            var pTitleStyle = ((SlideLayoutPart)openXmlPart).SlideMasterPart!.SlideMaster.TextStyles!
                .TitleStyle!;
            var masterTitleFonts = new IndentFonts(pTitleStyle);
            var masterTitleFont = masterTitleFonts.FontOrNull(indentLevel);
            if (this.HexFromName(masterTitleFont, out var masterTitleColor))
            {
                return masterTitleColor;
            }
        }
        else if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.Body)
        {
            var pBodyStyle = ((SlideLayoutPart)openXmlPart).SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!;
            var masterBodyFonts = new IndentFonts(pBodyStyle);
            var masterBodyFont = masterBodyFonts.FontOrNull(indentLevel);
            if (this.HexFromName(masterBodyFont, out var masterTitleColor))
            {
                return masterTitleColor;
            }
        }

        return null;
    }

    private string? SlideColorHexOrNull()
    {
        // Get basic shape and placeholder information
        var pShape = aText.Ancestors<P.Shape>().First();
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();

        if (pPlaceholderShape == null)
        {
            return null;
        }

        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aParagraph = aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();

        // Try to get color from layout shape
        var colorFromLayout = this.GetColorFromLayoutShape(pShape, indentLevel);
        if (colorFromLayout != null)
        {
            return colorFromLayout;
        }

        // Try to get color based on placeholder type
        return this.GetColorFromPlaceholderType(pPlaceholderShape, openXmlPart, indentLevel);
    }

    private string? GetColorFromLayoutShape(P.Shape pShape, int indentLevel)
    {
        var referencedLayoutPShape = this.ReferencedLayoutPShapeOrNull(pShape);

        // If no layout shape reference, try master shape
        if (referencedLayoutPShape == null)
        {
            return this.GetColorFromMasterShape(pShape, indentLevel);
        }

        // Check color from layout shape
        var layoutFonts = new IndentFonts(referencedLayoutPShape.TextBody!.ListStyle!);
        var layoutIndentFontOfSlide = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentFontOfSlide != null && this.HexFromName(layoutIndentFontOfSlide, out var layoutColorOfSlide))
        {
            return layoutColorOfSlide;
        }

        // Try master shape of layout if no color found
        return this.GetColorFromMasterShapeOfLayout(referencedLayoutPShape, indentLevel);
    }

    private string? GetColorFromMasterShape(P.Shape pShape, int indentLevel)
    {
        var referencedMasterPShape = this.ReferencedMasterPShapeOrNull(pShape);
        if (referencedMasterPShape == null)
        {
            return null;
        }

        var masterFontsOfSlide = new IndentFonts(referencedMasterPShape.TextBody!.ListStyle!);
        var masterIndentFontOfSlide = masterFontsOfSlide.FontOrNull(indentLevel);

        return this.HexFromName(masterIndentFontOfSlide, out var masterColorOfSlide)
            ? masterColorOfSlide
            : null;
    }

    private string? GetColorFromMasterShapeOfLayout(P.Shape layoutShape, int indentLevel)
    {
        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNull(layoutShape);
        if (refMasterPShapeOfLayout == null)
        {
            return null;
        }

        var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout.TextBody!.ListStyle!);
        var masterIndentFontOfLayout = masterFontsOfLayout.FontOrNull(indentLevel);

        return (masterIndentFontOfLayout != null &&
                this.HexFromName(masterIndentFontOfLayout, out var masterColorOfLayout))
            ? masterColorOfLayout
            : null;
    }

    private string? GetColorFromPlaceholderType(P.PlaceholderShape pPlaceholderShape, OpenXmlPart openXmlPart, int indentLevel)
    {
        if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.Title)
        {
            return this.GetColorFromTitlePlaceholder(openXmlPart, indentLevel);
        }

        return pPlaceholderShape.Type?.Value == P.PlaceholderValues.Body
            ? this.GetColorFromBodyPlaceholder(openXmlPart, indentLevel)
            :

            // No specific color handling for other placeholder types
            null;
    }

    private string? GetColorFromTitlePlaceholder(OpenXmlPart openXmlPart, int indentLevel)
    {
        var pTitleStyle = openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                .TitleStyle!,
            _ => ((SlideLayoutPart)openXmlPart).SlideMasterPart!.SlideMaster.TextStyles!
                .TitleStyle!
        };

        var masterTitleFonts = new IndentFonts(pTitleStyle);
        var masterTitleFont = masterTitleFonts.FontOrNull(indentLevel);

        return this.HexFromName(masterTitleFont, out var masterTitleColor)
            ? masterTitleColor
            : null;
    }

    private string? GetColorFromBodyPlaceholder(OpenXmlPart openXmlPart, int indentLevel)
    {
        var pBodyStyle = openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                .BodyStyle!,
            _ => ((SlideLayoutPart)openXmlPart).SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!
        };

        var masterBodyFonts = new IndentFonts(pBodyStyle);
        var masterBodyFont = masterBodyFonts.FontOrNull(indentLevel);

        return this.HexFromName(masterBodyFont, out var masterBodyColor)
            ? masterBodyColor
            : null;
    }

    private bool? SlideFontBoldFlagOrNull()
    {
        var aParagraph = aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
        var slidePShape = aText.Ancestors<P.Shape>().First();
        var slidePh = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return null;
        }

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNull(slidePShape);
        if (refLayoutPShapeOfSlide == null)
        {
            var refMasterPShape = this.ReferencedMasterPShapeOrNull(slidePShape);
            var fonts = new IndentFonts(refMasterPShape!.TextBody!.ListStyle!);

            return fonts.BoldFlagOrNull(indentLevel);
        }

        var layoutFonts = new IndentFonts(refLayoutPShapeOfSlide.TextBody!.ListStyle!);
        var layoutIndentColorType = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentColorType.HasValue)
        {
            return layoutIndentColorType.Value.IsBold;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNull(refLayoutPShapeOfSlide);
        if (refMasterPShapeOfLayout == null)
        {
            return null;
        }

        var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout.TextBody!.ListStyle!);
        var masterOfLayoutIndentColorType = masterFontsOfLayout.FontOrNull(indentLevel);
        if (masterOfLayoutIndentColorType.HasValue)
        {
            return masterOfLayoutIndentColorType.Value.IsBold;
        }

        return null;
    }

    private int? MasterFontSizeOrNull(P.Shape refLayoutPShapeOfSlide, int indentLevel)
    {
        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNull(refLayoutPShapeOfSlide);
        if (refMasterPShapeOfLayout == null)
        {
            return null;
        }

        var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout.TextBody!.ListStyle!);
        var masterOfLayoutIndentColorType = masterFontsOfLayout.FontOrNull(indentLevel);
        if (masterOfLayoutIndentColorType.HasValue && masterOfLayoutIndentColorType.Value.Size.HasValue)
        {
            return (int)masterOfLayoutIndentColorType.Value.Size!;
        }

        return null;
    }

    private A.LatinFont? SlideALatinFontOrNull(SlidePart sdkSlidePart)
    {
        var aParagraph = aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
        var pShape = aText.Ancestors<P.Shape>().FirstOrDefault();
        if (pShape == null)
        {
            return null;
        }

        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var refLayoutPShape = this.ReferencedLayoutPShapeOrNull(pShape);
        if (refLayoutPShape == null)
        {
            var refMasterPShape = this.ReferencedMasterPShapeOrNull(pShape);
            if (refMasterPShape == null)
            {
                if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.CenteredTitle)
                {
                    return sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!.TitleStyle!
                        .Level1ParagraphProperties!
                        .GetFirstChild<A.DefaultRunProperties>() !
                        .GetFirstChild<A.LatinFont>();
                }

                return null;
            }

            var fonts = new IndentFonts(refMasterPShape.TextBody!.ListStyle!);

            return fonts.ALatinFontOrNull(indentLevel);
        }

        var layoutFonts = new IndentFonts(refLayoutPShape.TextBody!.ListStyle!);
        var layoutIndentColorType = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentColorType.HasValue)
        {
            return layoutIndentColorType.Value.ALatinFont;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNull(refLayoutPShape);
        var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout!.TextBody!.ListStyle!);
        var masterOfLayoutIndentColorType = masterFontsOfLayout.FontOrNull(indentLevel);
        if (masterOfLayoutIndentColorType.HasValue)
        {
            return masterOfLayoutIndentColorType.Value.ALatinFont;
        }

        return null;
    }

    private A.LatinFont SlideMasterALatinFont()
    {
        var aParagraph = aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
        var pShape = aText.Ancestors<P.Shape>().First();
        var fonts = new IndentFonts(pShape.TextBody!.ListStyle!);

        return fonts.ALatinFontOrNull(indentLevel) !;
    }
}