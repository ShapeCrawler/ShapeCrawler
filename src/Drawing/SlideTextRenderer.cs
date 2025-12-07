using System;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Renders text content of shapes onto a canvas.
/// </summary>
internal sealed class SlideTextRenderer
{
    private const decimal DefaultFontSize = 12m;
    private readonly float defaultLineHeight;
    private readonly Func<string, double, SKColor> parseHexColor;
    private readonly float pointsToPixels;

    /// <summary>
    ///     Initializes a new instance of the <see cref="SlideTextRenderer"/> class.
    /// </summary>
    /// <param name="pointsToPixels">Points to pixels conversion factor.</param>
    /// <param name="parseHexColor">Delegate used to parse colors from hex.</param>
    internal SlideTextRenderer(float pointsToPixels, Func<string, double, SKColor> parseHexColor)
    {
        this.pointsToPixels = pointsToPixels;
        this.parseHexColor = parseHexColor;

        using var font = this.CreateFont(null);
        this.defaultLineHeight = font.Spacing;
    }

    /// <summary>
    ///     Renders the text of the specified shape.
    /// </summary>
    /// <param name="canvas">Canvas where the text is drawn.</param>
    /// <param name="shape">Shape that contains text.</param>
    internal void Render(SKCanvas canvas, IShape shape)
    {
        var textBox = shape.TextBox;
        if (textBox is null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(textBox.Text))
        {
            return;
        }

        var originX = ((float)shape.X + (float)textBox.LeftMargin) * this.pointsToPixels;
        var originY = ((float)shape.Y + (float)textBox.TopMargin) * this.pointsToPixels;

        var baseline = originY;
        foreach (var paragraph in textBox.Paragraphs)
        {
            this.RenderParagraph(canvas, paragraph, originX, ref baseline);
        }
    }

    private void RenderParagraph(SKCanvas canvas, IParagraph paragraph, float startX, ref float baseline)
    {
        var currentX = startX;
        var lineHeight = 0f;
        var hasLineContent = false;

        foreach (var portion in paragraph.Portions)
        {
            if (this.IsLineBreak(portion))
            {
                this.AdvanceLine(ref baseline, ref currentX, startX, ref lineHeight, ref hasLineContent);
                continue;
            }

            using var font = this.CreateFont(portion.Font);
            using var paint = this.CreatePaint(portion.Font);
            var metrics = font.Metrics;
            var drawY = baseline - metrics.Ascent;

            canvas.DrawText(portion.Text, currentX, drawY, SKTextAlign.Left, font, paint);

            currentX += font.MeasureText(portion.Text);
            lineHeight = Math.Max(lineHeight, font.Spacing);
            hasLineContent = true;
        }

        this.CompleteParagraph(ref baseline, ref lineHeight, ref hasLineContent);
    }

    private void CompleteParagraph(ref float baseline, ref float lineHeight, ref bool hasLineContent)
    {
        var heightToAdd = hasLineContent ? lineHeight : this.defaultLineHeight;
        baseline += heightToAdd <= 0 ? this.defaultLineHeight : heightToAdd;
        lineHeight = 0;
        hasLineContent = false;
    }

    private void AdvanceLine(
        ref float baseline,
        ref float currentX,
        float startX,
        ref float lineHeight,
        ref bool hasLineContent)
    {
        var heightToAdd = lineHeight > 0 ? lineHeight : this.defaultLineHeight;
        baseline += heightToAdd;
        currentX = startX;
        lineHeight = 0;
        hasLineContent = false;
    }

    private SKPaint CreatePaint(ITextPortionFont? font)
    {
        var paint = new SKPaint
        {
            IsAntialias = true,
            Style = SKPaintStyle.Fill,
            Color = this.GetPaintColor(font)
        };

        return paint;
    }

    private SKFont CreateFont(ITextPortionFont? font)
    {
        var fontStyle = this.GetFontStyle(font);
        var family = font?.LatinName;

        var typeface = string.IsNullOrWhiteSpace(family)
            ? SKTypeface.CreateDefault()
            : SKTypeface.FromFamilyName(family, fontStyle);

        return new SKFont(typeface) { Size = (float)(font?.Size ?? DefaultFontSize) * this.pointsToPixels };
    }

    private SKColor GetPaintColor(ITextPortionFont? font)
    {
        var hex = font?.Color.Hex;

        return string.IsNullOrWhiteSpace(hex)
            ? SKColors.Black
            : this.parseHexColor(hex!, 100);
    }

    private SKFontStyle GetFontStyle(ITextPortionFont? font)
    {
        var isBold = font?.IsBold == true;
        var isItalic = font?.IsItalic == true;

        if (isBold && isItalic)
        {
            return SKFontStyle.BoldItalic;
        }

        if (isBold)
        {
            return SKFontStyle.Bold;
        }

        if (isItalic)
        {
            return SKFontStyle.Italic;
        }

        return SKFontStyle.Normal;
    }

    private bool IsLineBreak(IParagraphPortion portion)
    {
        return portion.Text == Environment.NewLine;
    }
}
