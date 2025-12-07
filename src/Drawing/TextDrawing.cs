using System;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents text drawing.
/// </summary>
internal sealed class TextDrawing
{
    private const decimal DefaultFontSize = 12m;
    private readonly float defaultLineHeight;
    private readonly Func<string, double, SKColor> parseHexColor;

    internal TextDrawing(Func<string, double, SKColor> parseHexColor)
    {
        this.parseHexColor = parseHexColor;

        using var font = CreateFont(null);
        this.defaultLineHeight = font.Spacing;
    }

    internal void Render(SKCanvas canvas, IShape shape)
    {
        var textBox = shape.TextBox;
        if (textBox is null || string.IsNullOrWhiteSpace(textBox.Text))
        {
            return;
        }

        var originX = new Points(shape.X + textBox.LeftMargin).AsPixels();
        var originY = new Points(shape.Y + textBox.TopMargin).AsPixels();

        var baseline = (float)originY;
        foreach (var paragraph in textBox.Paragraphs)
        {
            this.RenderParagraph(canvas, paragraph, (float)originX, ref baseline);
        }
    }

    private void RenderParagraph(SKCanvas canvas, IParagraph paragraph, float startX, ref float baseline)
    {
        var currentX = startX;
        var lineHeight = 0f;
        var hasLineContent = false;

        foreach (var portion in paragraph.Portions)
        {
            if (IsLineBreak(portion))
            {
                this.AdvanceLine(ref baseline, ref currentX, startX, ref lineHeight, ref hasLineContent);
                continue;
            }

            using var font = CreateFont(portion.Font);
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

    private static SKFont CreateFont(ITextPortionFont? font)
    {
        var fontStyle = GetFontStyle(font);
        var family = font?.LatinName;

        var typeface = string.IsNullOrWhiteSpace(family)
            ? SKTypeface.CreateDefault()
            : SKTypeface.FromFamilyName(family, fontStyle);
        var size = new Points((font?.Size ?? DefaultFontSize)).AsPixels();

        return new SKFont(typeface) { Size = (float)size };
    }

    private SKColor GetPaintColor(ITextPortionFont? font)
    {
        var hex = font?.Color.Hex;

        return string.IsNullOrWhiteSpace(hex)
            ? SKColors.Black
            : this.parseHexColor(hex!, 100);
    }

    private static SKFontStyle GetFontStyle(ITextPortionFont? font)
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

    private static bool IsLineBreak(IParagraphPortion portion) => portion.Text == Environment.NewLine;
}
