using System.Linq;
using ShapeCrawler.Fonts;
using SkiaSharp;

namespace ShapeCrawler.Texts;

internal readonly ref struct Text(string content, ITextPortionFont font)
{
    internal decimal FontSize => font.Size;

    internal decimal Width
    {
        get
        {
            var fontFamily = string.IsNullOrEmpty(font.LatinName)
                ? "Calibri"
                : font.LatinName == "Calibri Light"
                    ? "Calibri" // for unknown reasons, SkiaSharp uses "Segoe UI" instead of "Calibri Light"
                    : font.LatinName;
            var weight = font.IsBold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
            var slant = font.IsItalic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
            var style = new SKFontStyle(weight, SKFontStyleWidth.Normal, slant);
            using var skFont = new SKFont
            {
                Size = (float)font.Size, Typeface = SKTypeface.FromFamilyName(fontFamily, style)
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
        var fontFamily = string.IsNullOrEmpty(font.LatinName) ? "Calibri" : font.LatinName;
        var weight = font.IsBold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal;
        var slant = font.IsItalic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright;
        var style = new SKFontStyle(weight, SKFontStyleWidth.Normal, slant);
        skFont.Typeface = SKTypeface.FromFamilyName(fontFamily, style);

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
        for (var i = 0; i < words.Count; i++)
        {
            var word = words[i];
            var wordWidth = skFont.MeasureText(word);
            
            // Handle word that fits on current line
            if (wordWidth <= rect.Right - wordX)
            {
                canvas.DrawText(word, wordX, wordY, SKTextAlign.Left, skFont, paint);
                wordX += wordWidth + spaceWidth;
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

        // Compensate for the scaling that will be applied later by PortionFontSize.ApplyNormAutofitScaling()
        var scaleFactor = GetNormAutofitScaleFactor(font);
        var compensatedSize = scaleFactor > 0 ? (decimal)skFont.Size / scaleFactor : (decimal)skFont.Size;
        
        font.Size = compensatedSize;
    }

    // Gets the scaling factor that will be applied by PortionFontSize.ApplyNormAutofitScaling()
    private static decimal GetNormAutofitScaleFactor(ITextPortionFont font)
    {
        // Access the underlying PortionFontSize to get the scaling factor
        if (font is TextPortionFont textPortionFont)
        {
            // Get the current font size without setting it
            var currentSize = textPortionFont.Size;
            
            // Set a test value and see how it gets scaled
            const decimal testSize = 100m;
            textPortionFont.Size = testSize;
            var scaledTestSize = textPortionFont.Size;
            
            // Restore the original size
            textPortionFont.Size = currentSize;
            
            // Calculate the scale factor
            return scaledTestSize / testSize;
        }
        
        return 1m; // No scaling if we can't determine it
    }

    // Resets the text layout coordinates when font size changes
    private static void ResetTextLayout(ref float wordX, ref float wordY, SKRect rect, SKFont font)
    {
        wordX = rect.Left;
        wordY = rect.Top + font.Size;
    }
}