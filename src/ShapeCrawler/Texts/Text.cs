using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Texts;

internal readonly ref struct Text
{
    private readonly string text;
    private readonly ITextPortionFont font;

    internal Text(string text, ITextPortionFont font)
    {
        this.text = text;
        this.font = font;
    }

    /// <summary>
    ///     Gets text width in pixels.
    /// </summary>
    internal decimal Width => this.GetWidth();

    private decimal GetWidth()
    {
        var fontFamily = this.font.LatinName == "Calibri Light" ? "Calibri" // for unknown reasons, SkiaSharp uses "Segoe UI" instead of "Calibri Light"
            : this.font.LatinName;
        var skFont = new SKFont
        {
            Size = new Points(this.font.Size).AsPixels(),
            Typeface = SKTypeface.FromFamilyName(fontFamily)
        };
        
        return (decimal)skFont.MeasureText(this.text);
    }
}