using System.Linq;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Texts;

internal readonly ref struct Text(string content, ITextPortionFont font)
{
    internal decimal FontSize => font.Size;

    internal decimal WidthPx
    {
        get
        {
            var fontFamily = font.LatinName == "Calibri Light"
                ? "Calibri" // for unknown reasons, SkiaSharp uses "Segoe UI" instead of "Calibri Light"
                : font.LatinName;
            var skFont = new SKFont
            {
                Size = new Points(font.Size).AsPixels(), Typeface = SKTypeface.FromFamilyName(fontFamily)
            };

            return (decimal)skFont.MeasureText(content);
        }
    }

    internal void Fit(decimal width, decimal height)
    {
        using var surface = SKSurface.Create(new SKImageInfo((int)width, (int)height));
        var canvas = surface.Canvas;
        using var paint = new SKPaint();
        paint.IsAntialias = true;

        using var skFont = new SKFont();
        skFont.Size = (float)font.Size;
        skFont.Typeface = SKTypeface.FromFamilyName(font.LatinName);

        const int defaultPaddingSize = 10;
        const int topBottomPadding = defaultPaddingSize * 2;
        var wordMaxY = height - topBottomPadding;

        var rect = new SKRect(
            defaultPaddingSize, 
            defaultPaddingSize, 
            (int)width - defaultPaddingSize,
            (int)height - defaultPaddingSize);

        var spaceWidth = skFont.MeasureText(" ");

        var wordX = rect.Left;
        var wordY = rect.Top + skFont.Size;

        var words = content.Split(' ').ToList();
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

                if ((decimal)wordY > wordMaxY)
                {
                    if (skFont.Size <= 5)
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

        font.Size = (decimal)skFont.Size;
    }
}