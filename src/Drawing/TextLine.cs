using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents a text line.
/// </summary>
internal sealed class TextLine(
    PixelTextPortion[] runs,
    TextHorizontalAlignment horizontalAlignment,
    float paraLeftMargin,
    float width,
    float height,
    float baselineOffset)
{
    private const decimal DefaultFontSize = 12m;

    internal PixelTextPortion[] Runs => runs;

    internal TextHorizontalAlignment HorizontalAlignment => horizontalAlignment;

    internal float ParaLeftMargin => paraLeftMargin;

    internal float Width => width;

    internal float Height => height;

    internal float BaselineOffset => baselineOffset;

    internal void Draw(SKCanvas canvas, float x, float y)
    {
        var baselineY = y + this.BaselineOffset;
        var currentX = x;

        foreach (var run in this.Runs)
        {
            using var font = this.CreateFont(run.Font);
            using var paint = this.CreatePaint(run.Font);

            canvas.DrawText(run.Text, currentX, baselineY, SKTextAlign.Left, font, paint);
            currentX += run.Width;
        }
    }

    private SKPaint CreatePaint(ITextPortionFont? font)
    {
        var paint = new SKPaint { IsAntialias = true, Style = SKPaintStyle.Fill, Color = this.GetPaintColor(font) };

        return paint;
    }

    private SKColor GetPaintColor(ITextPortionFont? font)
    {
        var hex = font?.Color.Hex;

        return string.IsNullOrWhiteSpace(hex)
            ? SKColors.Black
            : new Color(hex!).AsSkColor();
    }

    private SKFont CreateFont(ITextPortionFont? font)
    {
        var fontStyle = this.GetFontStyle(font);
        var family = font?.LatinName;

        var typeface = string.IsNullOrWhiteSpace(family)
            ? SKTypeface.CreateDefault()
            : SKTypeface.FromFamilyName(family, fontStyle);
        var size = new Points(font?.Size ?? DefaultFontSize).AsPixels();

        return new SKFont(typeface) { Size = (float)size };
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
}
