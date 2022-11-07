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
    SCAutoFitType AutofitType { get; }

    /// <summary>
    ///     Gets left margin of text frame in centimeters.
    /// </summary>
    double LeftMargin { get; }

    /// <summary>
    ///     Gets right margin of text frame in centimeters.
    /// </summary>
    double RightMargin { get; }

    /// <summary>
    ///     Gets top margin of text frame in centimeters.
    /// </summary>
    double TopMargin { get; }

    /// <summary>
    ///     Gets bottom margin of text frame in centimeters.
    /// </summary>
    double BottomMargin { get; }

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

    public SCAutoFitType AutofitType => this.GetAutoFitType();

    public double LeftMargin => this.GetLeftMargin();

    public double RightMargin => this.GetRightMargin();

    public double TopMargin => this.GetTopMargin();

    public double BottomMargin => this.GetBottomMargin();

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

    private double GetLeftMargin()
    {
        var bodyProperties = this.TextBodyElement.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.LeftInset;
        return ins is null ? SCConstants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private double GetRightMargin()
    {
        var bodyProperties = this.TextBodyElement.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.RightInset;
        return ins is null ? SCConstants.DefaultLeftAndRightMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private double GetTopMargin()
    {
        var bodyProperties = this.TextBodyElement.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.TopInset;
        return ins is null ? SCConstants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
    }

    private double GetBottomMargin()
    {
        var bodyProperties = this.TextBodyElement.GetFirstChild<A.BodyProperties>() !;
        var ins = bodyProperties.BottomInset;
        return ins is null ? SCConstants.DefaultTopAndBottomMargin : UnitConverter.EmuToCentimeter(ins.Value);
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

        if (this.AutofitType == SCAutoFitType.Shrink)
        {
            this.ShrinkText(newText, baseParagraph);
        }
        else if (this.AutofitType == SCAutoFitType.Resize)
        {
            this.ResizeShape(newText, baseParagraph);
        }

        baseParagraph.Text = newText;

        this.text.Reset();
    }

    private void ResizeShape(string newText, IParagraph baseParagraph)
    {
        var shape = this.TextFrameContainer.Shape;
        var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var paint = new SKPaint();
        var fontSize = font.Size;
        paint.TextSize = fontSize;
        paint.Typeface = SKTypeface.FromFamilyName(font.Name);
        paint.IsAntialias = true;

        var lMarginPixel = UnitConverter.CentimeterToPixel(this.LeftMargin);
        var rMarginPixel = UnitConverter.CentimeterToPixel(this.RightMargin);
        var tMarginPixel = UnitConverter.CentimeterToPixel(this.TopMargin);
        var bMarginPixel = UnitConverter.CentimeterToPixel(this.BottomMargin);

        var newTextRect = new SKRect();
        paint.MeasureText(newText, ref newTextRect);
        var newTextW = newTextRect.Width;
        var newTextH = paint.TextSize;
        var shapeTextBlockW = shape.Width - lMarginPixel - rMarginPixel;
        var shapeTextBlockH = shape.Height - tMarginPixel - bMarginPixel;
        if (newTextW > shapeTextBlockW)
        {
            var needRowsCount = newTextW / shapeTextBlockW;
            var intPart = (int)needRowsCount;
            var fractionalPart = needRowsCount - intPart;
            if (fractionalPart > 0)
            {
                intPart++;
            }

            var shapeNeedH = intPart * newTextH + tMarginPixel + bMarginPixel;
            if (shapeTextBlockH < shapeNeedH)
            {
                shape.Height = (int)shapeNeedH + tMarginPixel + bMarginPixel + tMarginPixel + bMarginPixel;

                // We should raise the shape up by the amount which is half of the increased offset.
                // PowerPoint does the same thing.
                var yOffset = (shapeNeedH - shapeTextBlockH) / 2;
                shape.Y -= (int)yOffset;
            }
        }
    }

    private void ShrinkText(string newText, IParagraph baseParagraph)
    {
        var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;
        var shape = this.TextFrameContainer.Shape;

        var fontSize = FontService.GetAdjustedFontSize(newText, font, shape);

        var paragraphInternal = (SCParagraph)baseParagraph;
        paragraphInternal.SetFontSize(fontSize);
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