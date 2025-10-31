using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents the autofit behavior that resizes text/shape to fit content.
/// </summary>
internal sealed class TextAutofit
{
    private readonly IParagraphCollection paragraphs;
    private readonly Func<AutofitType> getAutofitType;
    private readonly ShapeSize shapeSize;
    private readonly TextBoxMargins margins;
    private readonly Func<bool> getTextWrapped;
    private readonly OpenXmlElement textBody;

    internal TextAutofit(
        IParagraphCollection paragraphs,
        Func<AutofitType> getAutofitType,
        ShapeSize shapeSize,
        TextBoxMargins margins,
        Func<bool> getTextWrapped,
        OpenXmlElement textBody)
    {
        this.paragraphs = paragraphs;
        this.getAutofitType = getAutofitType;
        this.shapeSize = shapeSize;
        this.margins = margins;
        this.getTextWrapped = getTextWrapped;
        this.textBody = textBody;
    }

    /// <summary>
    ///     Applies autofit by resizing the parent shape on demand.
    /// </summary>
    internal void Apply()
    {
        if (this.getAutofitType() != AutofitType.Resize)
        {
            return;
        }

        var shapeWidthCapacity = this.shapeSize.Width - this.margins.Left - this.margins.Right;
        var shapeHeightCapacity = this.shapeSize.Height - this.margins.Top - this.margins.Bottom;

        decimal textHeight = 0;
        foreach (var paragraph in this.paragraphs)
        {
            var paragraphPortion = paragraph.Portions.OfType<TextParagraphPortion>();
            if (!paragraphPortion.Any())
            {
                continue;
            }

            var popularPortion = paragraphPortion.GroupBy(p => p.Font.Size)
                .OrderByDescending(x => x.Count())
                .First().First();
            var scFont = popularPortion.Font;

            var paragraphText = paragraph.Text.ToUpper();
            var paragraphTextWidth = new Text(paragraphText, scFont).Width;
            var paragraphTextHeight = scFont.Size;
            var requiredRowsCount = paragraphTextWidth / shapeWidthCapacity;
            var intRequiredRowsCount = (int)requiredRowsCount;
            var fractionalPart = requiredRowsCount - intRequiredRowsCount;
            if (fractionalPart > 0)
            {
                intRequiredRowsCount++;
            }

            textHeight += intRequiredRowsCount * (int)paragraphTextHeight;
        }

        this.UpdateHeight(textHeight, shapeHeightCapacity);
        if (!this.getTextWrapped())
        {
            this.UpdateWidth();
        }
    }

    /// <summary>
    ///     Shrinks font size to fit the text in the shape.
    /// </summary>
    internal void ShrinkFont(string newText)
    {
        var firstParagraph = this.paragraphs.First();
        var popularFont = firstParagraph.Portions.GroupBy(paraPortion => paraPortion.Font!.Size)
            .OrderByDescending(x => x.Count())
            .First().First().Font!;
        var text = new Text(newText, popularFont);
        text.Fit(this.shapeSize.Width, this.shapeSize.Height);
        firstParagraph.SetFontSize((int)text.FontSize);
    }

    private void UpdateWidth()
    {
        var longerText = this.paragraphs
            .Select(x => new { x.Text, x.Text.Length })
            .OrderByDescending(x => x.Length)
            .First().Text;

        var baseParagraph = this.paragraphs.First();
        var popularPortion = baseParagraph.Portions.OfType<TextParagraphPortion>().GroupBy(p => p.Font.Size)
            .OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var textWidth = new Text(longerText, font).Width;
        var leftMargin = this.margins.Left;
        var rightMargin = this.margins.Right;
        var newWidth =
            (int)(textWidth *
                  1.4M) // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi. 96/72 = 1.4 
            + leftMargin + rightMargin;
        this.shapeSize.Width = newWidth;
    }

    private void UpdateHeight(decimal textHeight, decimal shapeHeightCapacity)
    {
        var parentShape = this.textBody.Parent!;
        var requiredHeight = textHeight + this.margins.Top + this.margins.Bottom;
        var newHeight = requiredHeight + this.margins.Top + this.margins.Bottom + this.margins.Top + this.margins.Bottom;
        this.shapeSize.Height = newHeight;

        // Raise the shape up by the amount, which is half of the increased offset, like PowerPoint does it
        var position = new Position(parentShape);
        var yOffset = (requiredHeight - shapeHeightCapacity) / 2;
        position.Y -= yOffset;
    }
}