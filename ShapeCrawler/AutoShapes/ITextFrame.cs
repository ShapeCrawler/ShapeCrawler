// ReSharper disable CheckNamespace

using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Constants;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using ShapeCrawler.Statics;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

/// <summary>
///     Represents text frame.
/// </summary>
public interface ITextFrame
{
    /// <summary>
    ///     Gets collection of paragraphs.
    /// </summary>
    IParagraphCollection Paragraphs { get; }

    /// <summary>
    ///     Gets or sets text.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    ///     Gets Autofit type.
    /// </summary>
    SCAutoFitType AutoFitType { get; }

    /// <summary>
    ///     Gets left margin of text frame in centimeters.
    /// </summary>
    double LeftMargin { get; }

    /// <summary>
    ///     Gets right margin of text frame in centimeters.
    /// </summary>
    double RightMargin { get; }

    /// <summary>
    ///     Gets a value indicating whether text frame can be changed.
    /// </summary>
    bool CanChangeText();
}

internal class TextFrame : ITextFrame
{
    private readonly ResettableLazy<string> text;
    private readonly ResettableLazy<ParagraphCollection> paragraphs;

    internal TextFrame(ITextFrameContainer frameContainer, TypedOpenXmlCompositeElement textBodyElement)
    {
        this.TextFrameContainer = frameContainer;
        this.TextBodyElement = textBodyElement;
        this.text = new ResettableLazy<string>(this.GetText);
        this.paragraphs = new ResettableLazy<ParagraphCollection>(this.GetParagraphs);
    }

    public IParagraphCollection Paragraphs => this.paragraphs.Value;

    public string Text
    {
        get => this.text.Value;
        set => this.SetText(value);
    }

    public SCAutoFitType AutoFitType => this.GetAutoFitType();

    public double LeftMargin => this.GetLeftMargin();

    public double RightMargin => this.GetRightMargin();

    internal ITextFrameContainer TextFrameContainer { get; }

    internal OpenXmlCompositeElement TextBodyElement { get; }

    public bool CanChangeText()
    {
        var isField = this.Paragraphs.Any(paragraph => paragraph.Portions.Any(portion => portion.Field != null));
        var isFooter = this.TextFrameContainer.Shape.Placeholder?.Type == SCPlaceholderType.Footer;

        return !isField && !isFooter;
    }

    internal void Draw(SKCanvas slideCanvas, SKRect shapeRect)
    {
        throw new System.NotImplementedException();
    }

    private double GetRightMargin()
    {
        if (this.TextBodyElement is A.TextBody aTextBody)
        {
            var ins = aTextBody.BodyProperties!.RightInset;
            if (ins is null)
            {
                return SCConstants.DefaultLeftAndRightMargin;
            }

            return UnitConverter.EmuToCentimeter(ins.Value);
        }

        var pTextBody = (P.TextBody)this.TextBodyElement;
        var pTextBodyIns = pTextBody.BodyProperties!.RightInset;
        if (pTextBodyIns is null)
        {
            return SCConstants.DefaultLeftAndRightMargin;
        }

        return UnitConverter.EmuToCentimeter(pTextBodyIns.Value);
    }

    private double GetLeftMargin()
    {
        if (this.TextBodyElement is A.TextBody aTextBody)
        {
            var lIns = aTextBody.BodyProperties!.LeftInset;
            if (lIns is null)
            {
                return SCConstants.DefaultLeftAndRightMargin;
            }

            return UnitConverter.EmuToCentimeter(lIns.Value);
        }

        var pTextBody = (P.TextBody)this.TextBodyElement;
        var pTextBodyLIns = pTextBody.BodyProperties!.LeftInset;
        if (pTextBodyLIns is null)
        {
            return SCConstants.DefaultLeftAndRightMargin;
        }

        return UnitConverter.EmuToCentimeter(pTextBodyLIns.Value);
    }

    private ParagraphCollection GetParagraphs()
    {
        return new ParagraphCollection(this);
    }

    private SCAutoFitType GetAutoFitType()
    {
        if (this.TextBodyElement == null)
        {
            return SCAutoFitType.None;
        }

        var aBodyPr = this.TextBodyElement.GetFirstChild<A.BodyProperties>();
        if (aBodyPr!.GetFirstChild<A.NormalAutoFit>() != null)
        {
            return SCAutoFitType.Shrink;
        }

        if (aBodyPr.GetFirstChild<A.ShapeAutoFit>() != null)
        {
            return SCAutoFitType.Resize;
        }

        return SCAutoFitType.None;
    }

    private void SetText(string newText)
    {
        if (!this.CanChangeText())
        {
            throw new ShapeCrawlerException("Text can not be changed.");
        }

        var baseParagraph = this.Paragraphs.FirstOrDefault(p => p.Portions.Any());
        if (baseParagraph == null)
        {
            baseParagraph = this.Paragraphs.First();
            baseParagraph.AddPortion(newText);
        }

        var removingParagraphs = this.Paragraphs.Where(p => p != baseParagraph);
        this.Paragraphs.Remove(removingParagraphs);

        if (this.AutoFitType == SCAutoFitType.Shrink)
        {
            var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count())
                .First().First();
            var font = popularPortion.Font;
            var fontSize = popularPortion.Font.Size;
            var shape = this.TextFrameContainer.Shape;

            fontSize = FontService.GetAdjustedFontSize(newText, font, shape);

            var paragraphInternal = (SCParagraph)baseParagraph;
            paragraphInternal.SetFontSize(fontSize);
        }

        baseParagraph.Text = newText;

        // force the lazy property to refresh
        this.text.Reset();
    }

    private string GetText()
    {
        if (this.TextBodyElement == null)
        {
            return string.Empty;
        }

        var sb = new StringBuilder();
        sb.Append(this.Paragraphs[0].Text);

        // If the number of paragraphs more than one
        var numPr = this.Paragraphs.Count;
        var index = 1;
        while (index < numPr)
        {
            sb.AppendLine();
            sb.Append(this.Paragraphs[index].Text);

            index++;
        }

        return sb.ToString();
    }
}