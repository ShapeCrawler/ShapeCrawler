using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Fonts;
using ShapeCrawler.Paragraphs;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class ReferencedFont(ReferencedFontColor fontColor, A.Text aText)
{
    internal string? ColorHexOrNull()
    {
        return fontColor.ColorHexOrNull();
    }

    internal bool? BoldFlagOrNull()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (openXmlPart is SlidePart)
        {
            return this.SlideFontBoldFlagOrNull();
        }

        return null;
    }

    internal decimal? SizeOrNull()
    {
        var openXmlPart = aText.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var aParagraph = aText.Ancestors<A.Paragraph>().First();
        var indentLevel = new SCAParagraph(aParagraph).GetIndentLevel();
        var slidePShape = aText.Ancestors<P.Shape>().FirstOrDefault();
        if (slidePShape == null)
        {
            return null;
        }

        var pPlaceholderShape = slidePShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var refLayoutPShapeOfSlide = this.ReferencedLayoutPShapeOrNull(slidePShape);
        if (refLayoutPShapeOfSlide == null)
        {
            var refMasterPShape = this.ReferencedMasterPShapeOrNull(slidePShape);
            if (refMasterPShape != null)
            {
                var fonts = new IndentFonts(refMasterPShape.TextBody!.ListStyle!);
                var font = fonts.FontOrNull(indentLevel);
                if (font.HasValue)
                {
                    return (int)font.Value.Size! / 100m;
                }
            }

            var sdkSlidePart = (SlidePart)openXmlPart;
            var bodyStyleFonts =
                new IndentFonts(sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster!.TextStyles!.BodyStyle!);
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
            SlideLayoutPart or SlideMasterPart => this.SlideMasterALatinFont(),
            NotesSlidePart notesSlidePart => this.NotesSlideALatinFontOrNull(notesSlidePart),
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
            .GetFirstChild<P.PlaceholderShape>()!;
        var referencedLayoutPShape =
            new SCPShapeTree(slidePart.SlideLayoutPart!.SlideLayout!.CommonSlideData!.ShapeTree!).ReferencedPShapeOrNull(
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
            SlidePart slidePart => slidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster!.CommonSlideData!
                .ShapeTree!,
            _ => ((SlideLayoutPart)openXmlPart).SlideMasterPart!.SlideMaster!.CommonSlideData!
                .ShapeTree!
        };

        var referencedPShape = new SCPShapeTree(slideOrLayoutPShapeTree).ReferencedPShapeOrNull(pPlaceholderShape);

        return referencedPShape;
    }

    private bool? SlideFontBoldFlagOrNull()
    {
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
            if (refMasterPShape?.TextBody?.ListStyle == null)
            {
                return null;
            }

            var fonts = new IndentFonts(refMasterPShape.TextBody.ListStyle);

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
        if (masterOfLayoutIndentColorType is { Size: not null })
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
                    return sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster!.TextStyles!.TitleStyle!
                        .Level1ParagraphProperties!
                        .GetFirstChild<A.DefaultRunProperties>()!
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

        return fonts.ALatinFontOrNull(indentLevel)!;
    }

    private A.LatinFont? NotesSlideALatinFontOrNull(NotesSlidePart notesSlidePart)
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

        // NotesMaster doesn't have TextStyles like SlideMaster, so we fall back to the slide's master part
        var parentSlidePart = notesSlidePart.GetParentParts().OfType<SlidePart>().FirstOrDefault();
        if (parentSlidePart?.SlideLayoutPart?.SlideMasterPart != null)
        {
            var slideMasterFonts =
                new IndentFonts(parentSlidePart.SlideLayoutPart.SlideMasterPart.SlideMaster!.TextStyles!.BodyStyle!);
            var slideMasterIndentFont = slideMasterFonts.FontOrNull(indentLevel);
            if (slideMasterIndentFont.HasValue)
            {
                return slideMasterIndentFont.Value.ALatinFont;
            }
        }

        return null;
    }
}