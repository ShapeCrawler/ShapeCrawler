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
    
    internal TextHorizontalAlignment HorizontalAlignment => horizontalAlignment;

    internal float ParaLeftMargin => paraLeftMargin;

    internal float Width => width;

    internal float Height => height;

    private float BaselineOffset => baselineOffset;
    
    private PixelTextPortion[] Runs => runs;

    internal void Render(SKCanvas canvas, float x, float y)
    {
        var baselineY = y + this.BaselineOffset;
        var currentX = x;

        foreach (var run in this.Runs)
        {
            using var font = CreateFont(run.Font);
            using var paint = CreatePaint(run.Font);

            canvas.DrawText(run.Text, currentX, baselineY, SKTextAlign.Left, font, paint);
            currentX += run.Width;
        }
    }

    private static SKPaint CreatePaint(ITextPortionFont? font)
    {
        var paint = new SKPaint { IsAntialias = true, Style = SKPaintStyle.Fill, Color = GetPaintColor(font) };

        return paint;
    }

    private static SKColor GetPaintColor(ITextPortionFont? font)
    {
        var hex = font?.Color.Hex;

        return string.IsNullOrWhiteSpace(hex)
            ? SKColors.Black
            : new Color(hex!).AsSkColor();
    }

    private static SKFont CreateFont(ITextPortionFont? font)
    {
        var fontStyle = GetFontStyle(font);
        var family = font?.LatinName;

        var typeface = string.IsNullOrWhiteSpace(family)
            ? SKTypeface.CreateDefault()
            : SKTypeface.FromFamilyName(family, fontStyle);
        var size = new Points(font?.Size ?? DefaultFontSize).AsPixels();

        return new SKFont(typeface) { Size = (float)size };
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
}
