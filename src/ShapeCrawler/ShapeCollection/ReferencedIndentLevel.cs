using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Colors;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Fonts;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal readonly ref struct ReferencedIndentLevel
{
    private readonly OpenXmlPart sdkOpenXmlPart;
    private readonly A.Text aText;
    private readonly PresentationColor presColor;

    internal ReferencedIndentLevel(OpenXmlPart sdkOpenXmlPart, A.Text aText)
        : this(sdkOpenXmlPart, aText, new PresentationColor(sdkOpenXmlPart))
    {
    }

    private ReferencedIndentLevel(OpenXmlPart sdkOpenXmlPart, A.Text aText, PresentationColor presColor)
    {
        this.sdkOpenXmlPart = sdkOpenXmlPart;
        this.aText = aText;
        this.presColor = presColor;
    }

    internal string? ColorHexOrNull()
    {
        if (this.sdkOpenXmlPart is SlidePart)
        {
            return this.SlideColorHexOrNull();
        }

        if (this.sdkOpenXmlPart is SlideLayoutPart)
        {
            return this.LayoutColorHexOrNull();
        }

        throw new SCException("Not implemented.");
    }

    internal ColorType? ColorTypeOrNull()
    {
        if (this.sdkOpenXmlPart is SlidePart)
        {
            return this.SlideColorTypeOrNull();
        }

        return LayoutColorTypeOrNull();
    }

    internal bool? FontBoldFlagOrNull()
    {
        if (this.sdkOpenXmlPart is SlidePart)
        {
            return this.SlideFontBoldFlagOrNull();
        }

        return LayoutFontBoldFlagOrNull();
    }

    internal int? FontSizeOrNull() => this.SlideFontSizeOrNull();

    internal A.LatinFont? ALatinFontOrNull()
    {
        return this.sdkOpenXmlPart switch
        {
            SlidePart sdkSlidePart => this.SlideALatinFontOrNull(sdkSlidePart),
            SlideMasterPart => this.SlideMasterALatinFont(),
            _ => LayoutALatinFontOrNull()
        };
    }

    private static ColorType? LayoutColorTypeOrNull() => throw new System.NotImplementedException();

    private static bool? LayoutFontBoldFlagOrNull() => throw new System.NotImplementedException();

    private static A.LatinFont LayoutALatinFontOrNull() => throw new SCException("Not implemented.");

    private ColorType? MasterOfSlideIndentColorType(P.Shape slidePShape, int indentLevel)
    {
        var refMasterPShape = this.ReferencedMasterPShapeOrNullOf(slidePShape);
        if (refMasterPShape == null)
        {
            return null;
        }

        var fonts = new IndentFonts(refMasterPShape.TextBody!.ListStyle!);
        var colorType = fonts.ColorType(indentLevel);

        return colorType;
    }

    /// <summary>
    ///     Tries to get referenced Placeholder Shape located on Slide Layout.
    /// </summary>
    private P.Shape? ReferencedLayoutPShapeOrNull(P.Shape pShape)
    {
        if (this.sdkOpenXmlPart is not SlidePart slidePart)
        {
            return null;
        }

        var pPlaceholder = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>() !;
        var referencedLayoutPShape = new SPShapeTree(slidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!).ReferencedPShapeOrNull(pPlaceholder);

        return referencedLayoutPShape;
    }

    private P.Shape? ReferencedMasterPShapeOrNullOf(P.Shape pShape)
    {
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var slideOrLayoutPShapeTree = this.sdkOpenXmlPart switch
        {
            SlidePart slidePart => slidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!,
            _ => ((SlideLayoutPart)this.sdkOpenXmlPart).SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!
        };

        var referencedPShape = new SPShapeTree(slideOrLayoutPShapeTree).ReferencedPShapeOrNull(pPlaceholderShape);

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
            referencedShapeColorOrNull = this.presColor.ThemeColorHex(indentFont.Value.ASchemeColor.Val!.Value);
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
        var pShape = this.aText.Ancestors<P.Shape>().First();
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var referencedMasterPShape = this.ReferencedMasterPShapeOrNullOf(pShape);
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SAParagraph(aParagraph).IndentLevel();
        if (referencedMasterPShape != null)
        {
            var masterIndentFonts = new IndentFonts(referencedMasterPShape.TextBody!.ListStyle!);
            var masterIndentFont = masterIndentFonts.FontOrNull(indentLevel);
            if (masterIndentFont != null && this.HexFromName(masterIndentFont, out var masterColor))
            {
                return masterColor;
            }
        }

        if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.Title)
        {
            var pTitleStyle = ((SlideLayoutPart)this.sdkOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!
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
            var pBodyStyle = ((SlideLayoutPart)this.sdkOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!;
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
        var pShape = this.aText.Ancestors<P.Shape>().First();
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var referencedLayoutPShape = this.ReferencedLayoutPShapeOrNull(pShape);
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SAParagraph(aParagraph).IndentLevel();
        if (referencedLayoutPShape == null)
        {
            var referencedMasterPShape = this.ReferencedMasterPShapeOrNullOf(pShape);
            if (referencedMasterPShape == null)
            {
                return null;
            }

            var masterFontsOfSlide = new IndentFonts(referencedMasterPShape.TextBody!.ListStyle!);
            var masterIndentFontOfSlide = masterFontsOfSlide.FontOrNull(indentLevel);
            if (this.HexFromName(masterIndentFontOfSlide, out var masterColorOfSlide))
            {
                return masterColorOfSlide;
            }

            return null;
        }

        var layoutFonts = new IndentFonts(referencedLayoutPShape.TextBody!.ListStyle!);
        var layoutIndentFontOfSlide = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentFontOfSlide != null && this.HexFromName(layoutIndentFontOfSlide, out var layoutColorOfSlide))
        {
            return layoutColorOfSlide;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(referencedLayoutPShape);
        if (refMasterPShapeOfLayout != null)
        {
            var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout.TextBody!.ListStyle!);
            var masterIndentFontOfLayout = masterFontsOfLayout.FontOrNull(indentLevel);
            if (masterIndentFontOfLayout != null &&
                this.HexFromName(masterIndentFontOfLayout, out var masterColorOfLayout))
            {
                return masterColorOfLayout;
            }
        }

        if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.Title)
        {
            var pTitleStyle = this.sdkOpenXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                    .TitleStyle!,
                _ => ((SlideLayoutPart)this.sdkOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!
                    .TitleStyle!
            };
            var masterTitleFonts = new IndentFonts(pTitleStyle);
            var masterTitleFont = masterTitleFonts.FontOrNull(indentLevel);
            if (this.HexFromName(masterTitleFont, out var masterTitleColor))
            {
                return masterTitleColor;
            }
        }
        else if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.Body)
        {
            var pBodyStyle = this.sdkOpenXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                    .BodyStyle!,
                _ => ((SlideLayoutPart)this.sdkOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!
            };
            var masterBodyFonts = new IndentFonts(pBodyStyle);
            var masterBodyFont = masterBodyFonts.FontOrNull(indentLevel);
            if (this.HexFromName(masterBodyFont, out var masterTitleColor))
            {
                return masterTitleColor;
            }
        }

        return null;
    }

    private ColorType? SlideColorTypeOrNull()
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var slidePShape = this.aText.Ancestors<P.Shape>().First();
        var slidePh = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return null;
        }

        var referencedLayoutPShapeOrNull = this.ReferencedLayoutPShapeOrNull(slidePShape);
        var indentLevel = new SAParagraph(aParagraph).IndentLevel();
        if (referencedLayoutPShapeOrNull == null)
        {
            return this.MasterOfSlideIndentColorType(slidePShape, indentLevel);
        }

        var layoutFonts = new IndentFonts(referencedLayoutPShapeOrNull.TextBody!.ListStyle!);
        var layoutIndentColorType = layoutFonts.ColorType(indentLevel);
        if (layoutIndentColorType.HasValue)
        {
            return layoutIndentColorType;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(referencedLayoutPShapeOrNull);
        var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout!.TextBody!.ListStyle!);
        var masterOfLayoutIndentColorType = masterFontsOfLayout.ColorType(indentLevel);
        if (masterOfLayoutIndentColorType.HasValue)
        {
            return masterOfLayoutIndentColorType;
        }

        return null;
    }

    private bool? SlideFontBoldFlagOrNull()
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SAParagraph(aParagraph).IndentLevel();
        var slidePShape = this.aText.Ancestors<P.Shape>().First();
        var slidePh = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return null;
        }

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNull(slidePShape);
        if (refLayoutPShapeOfSlide == null)
        {
            var refMasterPShape = this.ReferencedMasterPShapeOrNullOf(slidePShape);
            var fonts = new IndentFonts(refMasterPShape!.TextBody!.ListStyle!);

            return fonts.BoldFlagOrNull(indentLevel);
        }

        var layoutFonts = new IndentFonts(refLayoutPShapeOfSlide.TextBody!.ListStyle!);
        var layoutIndentColorType = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentColorType.HasValue)
        {
            return layoutIndentColorType.Value.IsBold;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(refLayoutPShapeOfSlide);
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

    private int? SlideFontSizeOrNull()
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SAParagraph(aParagraph).IndentLevel();
        var slidePShape = this.aText.Ancestors<P.Shape>().FirstOrDefault();
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
            var refMasterPShape = this.ReferencedMasterPShapeOrNullOf(slidePShape);
            if (refMasterPShape != null)
            {
                var fonts = new IndentFonts(refMasterPShape!.TextBody!.ListStyle!);
                var font = fonts.FontOrNull(indentLevel);
                if (font.HasValue)
                {
                    return (int)font.Value.Size!;
                }
            }

            var sdkSlidePart = (SlidePart)this.sdkOpenXmlPart;
            var bodyStyleFonts =
                new IndentFonts(sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!);
            var bodyStyleFont = bodyStyleFonts.FontOrNull(indentLevel);
            if (bodyStyleFont.HasValue)
            {
                return (int)bodyStyleFont.Value.Size!;
            }

            return null;
        }

        var layoutFonts = new IndentFonts(refLayoutPShapeOfSlide.TextBody!.ListStyle!);
        var layoutIndentFont = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentFont.HasValue && layoutIndentFont.Value.Size.HasValue)
        {
            return (int)layoutIndentFont.Value.Size!;
        }

        return this.MasterFontSizeOrNull(refLayoutPShapeOfSlide, indentLevel);
    }

    private int? MasterFontSizeOrNull(P.Shape refLayoutPShapeOfSlide, int indentLevel)
    {
        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(refLayoutPShapeOfSlide);
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
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SAParagraph(aParagraph).IndentLevel();
        var pShape = this.aText.Ancestors<P.Shape>().FirstOrDefault();
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
            var refMasterPShape = this.ReferencedMasterPShapeOrNullOf(pShape);
            if (refMasterPShape == null)
            {
                if (pPlaceholderShape.Type!.Value == P.PlaceholderValues.CenteredTitle)
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

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(refLayoutPShape);
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
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SAParagraph(aParagraph).IndentLevel();
        var pShape = this.aText.Ancestors<P.Shape>().First();
        var fonts = new IndentFonts(pShape.TextBody!.ListStyle!);

        return fonts.ALatinFontOrNull(indentLevel) !;
    }
}