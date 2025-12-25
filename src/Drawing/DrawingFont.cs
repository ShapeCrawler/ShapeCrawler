using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

internal sealed class DrawingFont(ITextPortionFont? font)
{
    private const decimal DefaultFontSize = 12m;

    internal static float BaselineOffset(SKFont skFont)
    {
        var ascent = skFont.Metrics.Ascent;
        return ascent >= 0 ? 0 : -ascent;
    }

    internal SKFont AsSkFont()
    {
        var fontStyle = this.GetFontStyle();
        var family = font?.LatinName;

        var typeface = string.IsNullOrWhiteSpace(family)
            ? SKTypeface.CreateDefault()
            : SKTypeface.FromFamilyName(family, fontStyle);
        var size = new Points(font?.Size ?? DefaultFontSize).AsPixels();

        return new SKFont(typeface) { Size = (float)size };
    }

    internal SKPaint CreatePaint()
    {
        return new SKPaint { IsAntialias = true, Style = SKPaintStyle.Fill, Color = this.GetPaintColor() };
    }

    private SKColor GetPaintColor()
    {
        var hex = font?.Color.Hex;

        return string.IsNullOrWhiteSpace(hex)
            ? SKColors.Black
            : new Color(hex!).AsSkColor();
    }

    private SKFontStyle GetFontStyle()
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