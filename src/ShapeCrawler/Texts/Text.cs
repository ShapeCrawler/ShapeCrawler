using System;
using System.Linq;
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

    public decimal FontSize => this.font.Size;

    internal decimal PxWidth => this.GetPxWidth();

    internal void FitInto(decimal width, decimal height)
    {
        using var surface = SKSurface.Create(new SKImageInfo((int)width, (int)height));
        var canvas = surface.Canvas;

        using var paint = new SKPaint();
        paint.IsAntialias = true;

        using var skFont = new SKFont();
        skFont.Size = (float)this.font.Size;
        skFont.Typeface = SKTypeface.FromFamilyName(this.font.LatinName);

        const int defaultPaddingSize = 10;
        const int topBottomPadding = defaultPaddingSize * 2;
        var wordMaxY = height - topBottomPadding;

        var rect = new SKRect(defaultPaddingSize, defaultPaddingSize, (int)width - defaultPaddingSize, (int)height - defaultPaddingSize);

        var spaceWidth = skFont.MeasureText(" ");

        var wordX = rect.Left;
        var wordY = rect.Top + skFont.Size;

        var words = this.text.Split(' ').ToList();
        for (var i = 0; i < words.Count;)
        {
            var word = words[i];
            var wordWidth = skFont.MeasureText(word);
            if (wordWidth <= rect.Right - wordX)
            {
                canvas.DrawText(word, wordX, wordY, SKTextAlign.Left, skFont, paint);
                wordX += wordWidth + spaceWidth;
            }
            else
            {
                wordY += skFont.Spacing;
                wordX = rect.Left;

                if (wordY > (float)wordMaxY)
                {
                    if (skFont.Size <= 5) // Min reduce font size
                    {
                        break;
                    }

                    skFont.Size = --skFont.Size;
                    wordX = rect.Left;
                    wordY = rect.Top + skFont.Size;
                    i = -1;
                }
                else
                {
                    wordX += wordWidth + spaceWidth;
                    canvas.DrawText(word, wordX, wordY, SKTextAlign.Left, skFont, paint);
                }
            }

            i++;
        }

        const int dpi = 96;
        var points = Math.Round(skFont.Size * 72 / dpi, 0);

        this.font.Size = (decimal)points;
    }

    private decimal GetPxWidth()
    {
        var fontFamily = this.font.LatinName == "Calibri Light"
            ? "Calibri" // for unknown reasons, SkiaSharp uses "Segoe UI" instead of "Calibri Light"
            : this.font.LatinName;
        var skFont = new SKFont
        {
            Size = new Points(this.font.Size).AsPixels(), Typeface = SKTypeface.FromFamilyName(fontFamily)
        };

        return (decimal)skFont.MeasureText(this.text);
    }
}