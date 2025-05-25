using System.Linq;
using SkiaSharp;

namespace ShapeCrawler.Texts;

internal readonly ref struct Text(string content, ITextPortionFont font)
{
    internal decimal FontSize => font.Size;

    internal decimal Width
    {
        get
        {
            var fontFamily = font.LatinName == "Calibri Light"
                ? "Calibri" // for unknown reasons, SkiaSharp uses "Segoe UI" instead of "Calibri Light"
                : font.LatinName;
            var skFont = new SKFont
            {
                Size = (float)font.Size, Typeface = SKTypeface.FromFamilyName(fontFamily)
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
            
            // Handle word that fits on current line
            if (wordWidth <= rect.Right - wordX)
            {
                canvas.DrawText(word, wordX, wordY, SKTextAlign.Left, skFont, paint);
                wordX += wordWidth + spaceWidth;
                i++;
                continue;
            }
            
            // Move to next line
            wordY += skFont.Spacing;
            wordX = rect.Left;
            
            // Check if we've reached vertical limit
            if ((decimal)wordY > wordMaxY)
            {
                // Minimum font size reached, can't shrink further
                if (skFont.Size <= 5)
                {
                    break;
                }

                // Reduce font size and restart layout
                skFont.Size--;
                ResetTextLayout(ref wordX, ref wordY, rect, skFont);
                i = -1;
            }
            else
            {
                // Draw word at beginning of new line
                canvas.DrawText(word, wordX, wordY, SKTextAlign.Left, skFont, paint);
                wordX += wordWidth + spaceWidth;
            }
            
            i++;
        }

        font.Size = (decimal)skFont.Size;
    }

    // Resets the text layout coordinates when font size changes
    private static void ResetTextLayout(ref float wordX, ref float wordY, SKRect rect, SKFont font)
    {
        wordX = rect.Left;
        wordY = rect.Top + font.Size;
    }
}