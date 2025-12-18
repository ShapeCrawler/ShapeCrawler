using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Paragraphs;

/// <summary>
///     Resolves the effective horizontal alignment for a paragraph.
/// </summary>
internal sealed class ParagraphHorizontalAlignment
{
    private readonly A.Paragraph aParagraph;
    private readonly SCAParagraph scAParagraph;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ParagraphHorizontalAlignment"/> class.
    /// </summary>
    /// <param name="aParagraph">DrawingML paragraph.</param>
    internal ParagraphHorizontalAlignment(A.Paragraph aParagraph)
    {
        this.aParagraph = aParagraph;
        this.scAParagraph = new SCAParagraph(aParagraph);
    }

    /// <summary>
    ///     Returns the effective horizontal alignment or <see langword="null"/> when it is not defined anywhere.
    /// </summary>
    internal TextHorizontalAlignment? ValueOrNull()
    {
        var explicitAlignment = this.ExplicitAlignmentOrNull();
        if (explicitAlignment.HasValue)
        {
            return explicitAlignment.Value;
        }

        return this.ReferencedAlignmentOrNull();
    }

    private TextHorizontalAlignment? ExplicitAlignmentOrNull()
    {
        var aTextAlignmentType = this.aParagraph.ParagraphProperties?.Alignment?.Value;
        if (aTextAlignmentType is null)
        {
            return null;
        }

        return this.ToHorizontalAlignment(aTextAlignmentType.Value);
    }

    private TextHorizontalAlignment? ReferencedAlignmentOrNull()
    {
        var pShape = this.aParagraph.Ancestors<P.Shape>().FirstOrDefault();
        if (pShape?.TextBody == null)
        {
            return null;
        }

        var indentLevel = this.scAParagraph.GetIndentLevel();
        var textBodyAlignment = this.AlignmentFromIndentStylesOrNull(pShape.TextBody.ListStyle, indentLevel);
        if (textBodyAlignment.HasValue)
        {
            return textBodyAlignment.Value;
        }

        var pPlaceholderShape = pShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape is null)
        {
            return null;
        }

        var openXmlPart = this.aParagraph.Ancestors<OpenXmlPartRootElement>().FirstOrDefault()?.OpenXmlPart;
        if (openXmlPart is null)
        {
            return null;
        }

        var layoutAlignment = this.AlignmentFromReferencedLayoutOrNull(openXmlPart, pPlaceholderShape, indentLevel);
        if (layoutAlignment.HasValue)
        {
            return layoutAlignment.Value;
        }

        var masterShapeAlignment = this.AlignmentFromReferencedMasterShapeOrNull(openXmlPart, pPlaceholderShape, indentLevel);
        if (masterShapeAlignment.HasValue)
        {
            return masterShapeAlignment.Value;
        }

        return this.AlignmentFromSlideMasterTextStylesOrNull(openXmlPart, pPlaceholderShape, indentLevel);
    }

    private TextHorizontalAlignment ToHorizontalAlignment(A.TextAlignmentTypeValues value)
    {
        if (value == A.TextAlignmentTypeValues.Center)
        {
            return TextHorizontalAlignment.Center;
        }

        if (value == A.TextAlignmentTypeValues.Right)
        {
            return TextHorizontalAlignment.Right;
        }

        if (value == A.TextAlignmentTypeValues.Justified)
        {
            return TextHorizontalAlignment.Justify;
        }

        return TextHorizontalAlignment.Left;
    }

    private TextHorizontalAlignment? AlignmentFromReferencedLayoutOrNull(
        OpenXmlPart openXmlPart,
        P.PlaceholderShape pPlaceholderShape,
        int indentLevel)
    {
        if (openXmlPart is not SlidePart slidePart)
        {
            return null;
        }

        var layoutShapeTree = slidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!;
        var referencedLayoutShape = new SCPShapeTree(layoutShapeTree).ReferencedPShapeOrNull(pPlaceholderShape);
        if (referencedLayoutShape?.TextBody?.ListStyle is null)
        {
            return null;
        }

        return this.AlignmentFromIndentStylesOrNull(referencedLayoutShape.TextBody.ListStyle, indentLevel);
    }

    private TextHorizontalAlignment? AlignmentFromReferencedMasterShapeOrNull(
        OpenXmlPart openXmlPart,
        P.PlaceholderShape pPlaceholderShape,
        int indentLevel)
    {
        P.ShapeTree? shapeTree = null;
        if (openXmlPart is SlidePart slidePart)
        {
            shapeTree = slidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!.ShapeTree!;
        }
        else if (openXmlPart is SlideLayoutPart slideLayoutPart)
        {
            shapeTree = slideLayoutPart.SlideMasterPart!.SlideMaster.CommonSlideData!.ShapeTree!;
        }
        else if (openXmlPart is SlideMasterPart slideMasterPart)
        {
            shapeTree = slideMasterPart.SlideMaster.CommonSlideData!.ShapeTree!;
        }

        if (shapeTree is null)
        {
            return null;
        }

        var referencedMasterShape = new SCPShapeTree(shapeTree).ReferencedPShapeOrNull(pPlaceholderShape);
        if (referencedMasterShape?.TextBody?.ListStyle is null)
        {
            return null;
        }

        return this.AlignmentFromIndentStylesOrNull(referencedMasterShape.TextBody.ListStyle, indentLevel);
    }

    private TextHorizontalAlignment? AlignmentFromSlideMasterTextStylesOrNull(
        OpenXmlPart openXmlPart,
        P.PlaceholderShape pPlaceholderShape,
        int indentLevel)
    {
        var slideMasterPart = this.SlideMasterPartOrNull(openXmlPart);
        var textStyles = slideMasterPart?.SlideMaster.TextStyles;
        if (textStyles is null)
        {
            return null;
        }

        OpenXmlCompositeElement? styles;
        var placeholderType = pPlaceholderShape.Type?.Value;
        if (placeholderType == P.PlaceholderValues.Title || placeholderType == P.PlaceholderValues.CenteredTitle)
        {
            styles = textStyles.TitleStyle;
        }
        else if (placeholderType == P.PlaceholderValues.Body)
        {
            styles = textStyles.BodyStyle;
        }
        else
        {
            styles = textStyles.OtherStyle;
        }

        return this.AlignmentFromIndentStylesOrNull(styles, indentLevel);
    }

    private SlideMasterPart? SlideMasterPartOrNull(OpenXmlPart openXmlPart)
    {
        if (openXmlPart is SlidePart slidePart)
        {
            return slidePart.SlideLayoutPart?.SlideMasterPart;
        }

        if (openXmlPart is SlideLayoutPart slideLayoutPart)
        {
            return slideLayoutPart.SlideMasterPart;
        }

        return openXmlPart is SlideMasterPart slideMasterPart ? slideMasterPart : null;
    }

    private TextHorizontalAlignment? AlignmentFromIndentStylesOrNull(
        OpenXmlCompositeElement? openXmlCompositeElement,
        int indentLevel)
    {
        if (openXmlCompositeElement is null)
        {
            return null;
        }

        foreach (var levelProperties in openXmlCompositeElement.Elements()
                     .Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal)))
        {
            var level = this.ExtractLevelNumberOrZero(levelProperties.LocalName);
            if (level != indentLevel)
            {
                continue;
            }

            var alignment = this.AlignmentFromAttributesOrNull(levelProperties);
            if (alignment.HasValue)
            {
                return alignment.Value;
            }
        }

        return null;
    }

    private int ExtractLevelNumberOrZero(string localName)
    {
        if (localName.Length < 4)
        {
            return 0;
        }

        var levelChar = localName[3];
        return levelChar >= '0' && levelChar <= '9' ? levelChar - '0' : 0;
    }

    private TextHorizontalAlignment? AlignmentFromAttributesOrNull(OpenXmlElement levelParagraphProperties)
    {
        var rawAlignment = levelParagraphProperties.GetAttributes().FirstOrDefault(a => a.LocalName == "algn").Value;
        if (string.IsNullOrWhiteSpace(rawAlignment))
        {
            return null;
        }

        var algn = rawAlignment!.Trim().ToLowerInvariant();

        if (algn == "l" || algn == "left")
        {
            return TextHorizontalAlignment.Left;
        }

        if (algn == "ctr" || algn == "center")
        {
            return TextHorizontalAlignment.Center;
        }

        if (algn == "r" || algn == "right")
        {
            return TextHorizontalAlignment.Right;
        }

        if (algn == "just" || algn == "justlow" || algn == "dist" || algn == "thaidist")
        {
            return TextHorizontalAlignment.Justify;
        }

        return null;
    }
}