using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;

namespace ShapeCrawler.Texts;

/// <summary>
///     Represents an autofit behavior that resizes text/shape to fit content.
/// </summary>
internal sealed class TextAutofit(
    IParagraphCollection paragraphs,
    Func<AutofitType> getAutofitType,
    ShapeSize shapeSize,
    TextBoxMargins margins,
    Func<bool> getTextWrapped,
    OpenXmlElement textBody)
{
    /// <summary>
    ///     Applies to autofit by resizing the parent shape on demand.
    /// </summary>
    internal void Apply()
    {
        if (getAutofitType() != AutofitType.Resize)
        {
            return;
        }

        var shapeWidthCapacity = shapeSize.Width - margins.Left - margins.Right;
        var shapeHeightCapacity = shapeSize.Height - margins.Top - margins.Bottom;

        decimal textHeight = 0;
        foreach (var paragraph in paragraphs)
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
        if (!getTextWrapped())
        {
            this.UpdateWidth();
        }
    }

    /// <summary>
    ///     Shrinks font size to fit the text in the shape.
    /// </summary>
    internal void ShrinkFont(string newText)
    {
        var firstParagraph = paragraphs.First();
        var popularFont = firstParagraph.Portions.GroupBy(paraPortion => paraPortion.Font!.Size)
            .OrderByDescending(x => x.Count())
            .First().First().Font!;
        var text = new Text(newText, popularFont);
        text.Fit(shapeSize.Width, shapeSize.Height);
        firstParagraph.SetFontSize((int)text.FontSize);
    }

    private void UpdateWidth()
    {
        var longerText = paragraphs
            .Select(x => new { x.Text, x.Text.Length })
            .OrderByDescending(x => x.Length)
            .First().Text;

        var baseParagraph = paragraphs.First();
        var popularPortion = baseParagraph.Portions.OfType<TextParagraphPortion>().GroupBy(p => p.Font.Size)
            .OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var textWidth = new Text(longerText, font).Width;
        var leftMargin = margins.Left;
        var rightMargin = margins.Right;
        var newWidth =
            (int)(textWidth *
                  1.4M) // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi. 96/72 = 1.4 
            + leftMargin + rightMargin;
        shapeSize.Width = newWidth;
    }

    private void UpdateHeight(decimal textHeight, decimal shapeHeightCapacity)
    {
        var parentShape = textBody.Parent!;
        var requiredHeight = textHeight + margins.Top + margins.Bottom;
        var newHeight = requiredHeight + margins.Top + margins.Bottom + margins.Top +
                        margins.Bottom;
        shapeSize.Height = newHeight;

        // Raise the shape up by the amount, which is half of the increased offset, like PowerPoint does it
        var position = new Position(parentShape);
        var yOffset = (requiredHeight - shapeHeightCapacity) / 2;
        position.Y -= yOffset;
    }
}