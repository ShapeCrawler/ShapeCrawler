using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Colors;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Fonts;
using ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal readonly record struct ReferencedIndent
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly A.Text aText;
    private readonly PresentationColor presColor;

    internal ReferencedIndent(TypedOpenXmlPart sdkTypedOpenXmlPart, A.Text aText)
        : this(sdkTypedOpenXmlPart, aText, new PresentationColor(sdkTypedOpenXmlPart))
    {
    }

    private ReferencedIndent(TypedOpenXmlPart sdkTypedOpenXmlPart, A.Text aText, PresentationColor presColor)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.aText = aText;
        this.presColor = presColor;
    }

    #region APIs

    /// <summary>
    ///     Color's hexadecimal representation from Referenced Layout or Master Shape for specified Slide Shape.
    /// </summary>
    internal string? ColorHexOrNull()
    {
        if (this.sdkTypedOpenXmlPart is SlidePart)
        {
            return this.SlideColorHexOrNull();
        }

        if (this.sdkTypedOpenXmlPart is SlideLayoutPart)
        {
            return this.LayoutColorHexOrNull();
        }

        throw new SCException("Not implemented.");
    }

    internal ColorType? ColorTypeOrNull()
    {
        if (this.sdkTypedOpenXmlPart is SlidePart)
        {
            return this.SlideColorTypeOrNull();
        }

        return this.LayoutColorTypeOrNull();
    }

    internal bool? FontBoldFlagOrNull()
    {
        if (this.sdkTypedOpenXmlPart is SlidePart)
        {
            return this.SlideFontBoldFlagOrNull();
        }

        return this.LayoutFontBoldFlagOrNull();
    }

    internal A.LatinFont? ALatinFontOrNull()
    {
        if (this.sdkTypedOpenXmlPart is SlidePart sdkSlidePart)
        {
            return this.SlideALatinFontOrNull(sdkSlidePart);
        }

        return this.LayoutALatinFontOrNull();
    }

    #endregion APIs

    /// <summary>
    ///     Tries to get color type from Referenced Master Placeholder of Slide Shape.
    /// </summary>
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
    private P.Shape? ReferencedLayoutPShapeOrNullOf(P.Shape sourcePShape)
    {
        if (this.sdkTypedOpenXmlPart is not SlidePart sdkSlidePart)
        {
            return null;
        }

        var slidePh = sourcePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>() !;

        var layoutPShapes =
            sdkSlidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!.Elements<P.Shape>();

        if (ReferencedPShape(layoutPShapes, slidePh, out var referencedPShape))
        {
            return referencedPShape;
        }

        return null;
    }

    private P.Shape? ReferencedMasterPShapeOrNullOf(P.Shape pShape)
    {
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var masterPShapes = this.sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>(),
            _ => ((SlideLayoutPart)this.sdkTypedOpenXmlPart).SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>()
        };

        if (ReferencedPShape(masterPShapes, pPlaceholderShape, out var referencedPShape))
        {
            return referencedPShape;
        }

        return null;
    }

    private static bool ReferencedPShape(
        IEnumerable<P.Shape> layoutPShapes,
        P.PlaceholderShape slidePh,
        out P.Shape? referencedShape)
    {
        foreach (var layoutPShape in layoutPShapes)
        {
            var layoutPh = layoutPShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>();
            if (layoutPh == null)
            {
                continue;
            }

            if (slidePh.Index is not null && layoutPh.Index is not null &&
                slidePh.Index == layoutPh.Index)
            {
                referencedShape = layoutPShape;
                return true;
            }

            if (slidePh.Type == null || layoutPh.Type == null)
            {
                referencedShape = layoutPShape;
                return true;
            }

            if (slidePh.Type == P.PlaceholderValues.Body &&
                slidePh.Index is not null && layoutPh.Index is not null)
            {
                if (slidePh.Index == layoutPh.Index)
                {
                    referencedShape = layoutPShape;
                    return true;
                }
            }

            if (slidePh.Type == P.PlaceholderValues.Title && layoutPh.Type == P.PlaceholderValues.Title)
            {
                referencedShape = layoutPShape;
                return true;
            }
            
            if (slidePh.Type == P.PlaceholderValues.CenteredTitle && layoutPh.Type == P.PlaceholderValues.CenteredTitle)
            {
                referencedShape = layoutPShape;
                return true;
            }
        }

        var byType = layoutPShapes.FirstOrDefault(layoutPShape =>
            layoutPShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == slidePh.Type?.Value);
        if (byType != null)
        {
            referencedShape = byType;
            return true;
        }

        referencedShape = null;
        return false;
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
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new AParagraphWrap(aParagraph).IndentLevel();
        var pShape = this.aText.Ancestors<P.Shape>().First();
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(pShape);
        if (refMasterPShapeOfLayout != null)
        {
            var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout.TextBody!.ListStyle!);
            var masterIndentFontOfLayout = masterFontsOfLayout.FontOrNull(indentLevel);
            if (masterIndentFontOfLayout != null
                && this.HexFromName(masterIndentFontOfLayout, out var masterColorOfLayout))
            {
                return masterColorOfLayout;
            }
        }

        if (pPlaceholderShape.Type?.Value == P.PlaceholderValues.Title)
        {
            var pTitleStyle = this.sdkTypedOpenXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                    .TitleStyle!,
                _ => ((SlideLayoutPart)this.sdkTypedOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!
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
            var pBodyStyle = this.sdkTypedOpenXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                    .BodyStyle!,
                _ => ((SlideLayoutPart)this.sdkTypedOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!
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

    private string? SlideColorHexOrNull()
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new AParagraphWrap(aParagraph).IndentLevel();
        var pShape = this.aText.Ancestors<P.Shape>().First();
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNullOf(pShape);
        if (refLayoutPShapeOfSlide == null)
        {
            var refMasterPShapeOfSlide = this.ReferencedMasterPShapeOrNullOf(pShape);
            if (refMasterPShapeOfSlide == null)
            {
                return null;
            }

            var masterFontsOfSlide = new IndentFonts(refMasterPShapeOfSlide!.TextBody!.ListStyle!);
            var masterIndentFontOfSlide = masterFontsOfSlide.FontOrNull(indentLevel);
            if (this.HexFromName(masterIndentFontOfSlide, out var masterColorOfSlide))
            {
                return masterColorOfSlide;
            }

            return null;
        }

        var layoutFonts = new IndentFonts(refLayoutPShapeOfSlide.TextBody!.ListStyle!);
        var layoutIndentFontOfSlide = layoutFonts.FontOrNull(indentLevel);
        if (layoutIndentFontOfSlide != null && this.HexFromName(layoutIndentFontOfSlide, out var layoutColorOfSlide))
        {
            return layoutColorOfSlide;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(refLayoutPShapeOfSlide);
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
            var pTitleStyle = this.sdkTypedOpenXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                    .TitleStyle!,
                _ => ((SlideLayoutPart)this.sdkTypedOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!
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
            var pBodyStyle = this.sdkTypedOpenXmlPart switch
            {
                SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!
                    .BodyStyle!,
                _ => ((SlideLayoutPart)this.sdkTypedOpenXmlPart).SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!
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

    private ColorType? LayoutColorTypeOrNull()
    {
        throw new System.NotImplementedException();
    }

    private ColorType? SlideColorTypeOrNull()
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new AParagraphWrap(aParagraph).IndentLevel();
        var slidePShape = this.aText.Ancestors<P.Shape>().First();
        var slidePh = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return null;
        }

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNullOf(slidePShape);
        if (refLayoutPShapeOfSlide == null)
        {
            return this.MasterOfSlideIndentColorType(slidePShape, indentLevel);
        }

        var layoutFonts = new IndentFonts(refLayoutPShapeOfSlide.TextBody!.ListStyle!);
        var layoutIndentColorType = layoutFonts.ColorType(indentLevel);
        if (layoutIndentColorType.HasValue)
        {
            return layoutIndentColorType;
        }

        var refMasterPShapeOfLayout = this.ReferencedMasterPShapeOrNullOf(refLayoutPShapeOfSlide);
        var masterFontsOfLayout = new IndentFonts(refMasterPShapeOfLayout!.TextBody!.ListStyle!);
        var masterOfLayoutIndentColorType = masterFontsOfLayout.ColorType(indentLevel);
        if (masterOfLayoutIndentColorType.HasValue)
        {
            return masterOfLayoutIndentColorType;
        }

        return null;
    }

    private bool? LayoutFontBoldFlagOrNull()
    {
        throw new System.NotImplementedException();
    }

    private bool? SlideFontBoldFlagOrNull()
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new AParagraphWrap(aParagraph).IndentLevel();
        var slidePShape = this.aText.Ancestors<P.Shape>().First();
        var slidePh = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (slidePh == null)
        {
            return null;
        }

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNullOf(slidePShape);
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
    
    internal int? FontSizeOrNull()
    {
        return this.SlideFontSizeOrNull();
    }
    
    private int? SlideFontSizeOrNull()
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new AParagraphWrap(aParagraph).IndentLevel();
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

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNullOf(slidePShape);
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

            var sdkSlidePart = (SlidePart)this.sdkTypedOpenXmlPart;
            var bodyStyleFonts = new IndentFonts(sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!.BodyStyle!);
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

    private A.LatinFont LayoutALatinFontOrNull()
    {
        throw new SCException("Not implemented.");
    }

    private A.LatinFont? SlideALatinFontOrNull(SlidePart sdkSlidePart)
    {
        var aParagraph = this.aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new AParagraphWrap(aParagraph).IndentLevel();
        var pShape = this.aText.Ancestors<P.Shape>().First();
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var refLayoutPShape = this.ReferencedLayoutPShapeOrNullOf(pShape);
        if (refLayoutPShape == null)
        {
            var refMasterPShape = this.ReferencedMasterPShapeOrNullOf(pShape);
            if (refMasterPShape == null)
            {
                if (pPlaceholderShape.Type!.Value == P.PlaceholderValues.CenteredTitle)
                {
                    return sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.TextStyles!.TitleStyle!.Level1ParagraphProperties!
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
}