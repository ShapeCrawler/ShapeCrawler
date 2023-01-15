using System;
using System.Linq;
using ShapeCrawler.Constants;
using SkiaSharp;

namespace ShapeCrawler.Services;

internal static class FontService
{
    internal static int GetAdjustedFontSize(string text, IFont font, SCShape scShape)
    {
        var surface = SKSurface.Create(new SKImageInfo(scShape.Width, scShape.Height));
        var canvas = surface.Canvas;

        var paint = new SKPaint();
        var fontSize = font.Size;
        paint.TextSize = fontSize;
        paint.Typeface = SKTypeface.FromFamilyName(font.LatinName);
        paint.IsAntialias = true;
        const int defaultPaddingSize = 10;
        const int topBottomPadding = defaultPaddingSize * 2;
        var wordMaxY = scShape.Height - topBottomPadding;

        var rect = new SKRect(defaultPaddingSize, defaultPaddingSize, scShape.Width - defaultPaddingSize, scShape.Height - defaultPaddingSize);

        var spaceWidth = paint.MeasureText(" ");

        var wordX = rect.Left;
        var wordY = rect.Top + paint.TextSize;

        var words = text.Split(' ').ToList();
        for (var i = 0; i < words.Count;)
        {
            var word = words[i];
            var wordWidth = paint.MeasureText(word);
            if (wordWidth <= rect.Right - wordX)
            {
                canvas.DrawText(word, wordX, wordY, paint);
                wordX += wordWidth + spaceWidth;
            }
            else
            {
                wordY += paint.FontSpacing;
                wordX = rect.Left;

                if (wordY > wordMaxY)
                {
                    if (paint.TextSize == SCConstants.MinReduceFontSize)
                    {
                        break;
                    }

                    paint.TextSize = --paint.TextSize;
                    wordX = rect.Left;
                    wordY = rect.Top + paint.TextSize;
                    i = -1;
                }
                else
                {
                    wordX += wordWidth + spaceWidth;
                    canvas.DrawText(word, wordX, wordY, paint);
                }
            }

            i++;
        }

        const int dpi = 96;
        var points = Math.Round(paint.TextSize * 72 / dpi, 0);

        return (int)points;
    }
}