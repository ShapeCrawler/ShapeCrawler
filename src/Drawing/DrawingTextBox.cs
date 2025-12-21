using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents text drawing.
/// </summary>
internal sealed class DrawingTextBox : TextBox
{
    private const decimal DefaultFontSize = 12m;
    private readonly float defaultLineHeight;
    private readonly float defaultBaselineOffset;

    internal DrawingTextBox(TextBoxMargins margins, OpenXmlElement textBody)
        : base(margins, textBody)
    {
        using var font = CreateFont(null);
        defaultLineHeight = font.Spacing;
        defaultBaselineOffset = GetBaselineOffset(font);
    }

    internal void Render(SKCanvas canvas, decimal parentShapeX, decimal parentShapeY, decimal parentShapeWidth, decimal parentShapeHeight)
    {
        if (string.IsNullOrWhiteSpace(Text))
        {
            return;
        }

        var originX = (float)new Points(parentShapeX + LeftMargin).AsPixels();
        var originY = (float)new Points(parentShapeY + TopMargin).AsPixels();
        var availableWidth = GetAvailableWidth(parentShapeWidth);
        var availableHeight = GetAvailableHeight(parentShapeHeight);

        var wrap = this.TextWrapped && availableWidth > 0;
        var lines = LayoutLines(availableWidth, wrap);
        var textBlockHeight = lines.Sum(l => l.Height);

        var verticalOffset = GetVerticalOffset(VerticalAlignment, availableHeight, textBlockHeight);
        var lineTop = originY + verticalOffset;

        foreach (var line in lines)
        {
            var horizontalOffset = GetHorizontalOffset(line.HorizontalAlignment, availableWidth - line.ParaLeftMargin, line.Width);
            var startX = originX + line.ParaLeftMargin + horizontalOffset;

            RenderLine(canvas, line, startX, lineTop);
            lineTop += line.Height;
        }
    }

    private static float GetVerticalOffset(TextVerticalAlignment alignment, float availableHeight, float contentHeight)
    {
        if (availableHeight <= 0)
        {
            return 0;
        }

        var freeSpace = availableHeight - contentHeight;
        return alignment switch
        {
            TextVerticalAlignment.Top => 0,
            TextVerticalAlignment.Middle => freeSpace / 2,
            TextVerticalAlignment.Bottom => freeSpace,
            _ => 0
        };
    }

    private static float GetHorizontalOffset(TextHorizontalAlignment alignment, float availableWidth, float lineWidth)
    {
        if (availableWidth <= 0)
        {
            return 0;
        }

        var freeSpace = availableWidth - lineWidth;
        return alignment switch
        {
            TextHorizontalAlignment.Left => 0,
            TextHorizontalAlignment.Center => freeSpace / 2,
            TextHorizontalAlignment.Right => freeSpace,
            _ => 0 // Treat Justify and unknown values as Left for MVP.
        };
    }

    private static bool IsWhitespace(string value) => string.IsNullOrEmpty(value) || value.All(char.IsWhiteSpace);

    private static IEnumerable<string> SplitToTokens(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            yield break;
        }

        var start = 0;
        while (start < text.Length)
        {
            var isWhitespace = char.IsWhiteSpace(text[start]);
            var index = start + 1;
            while (index < text.Length && char.IsWhiteSpace(text[index]) == isWhitespace)
            {
                index++;
            }

            yield return text[start..index];
            start = index;
        }
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

    private static bool IsLineBreak(IParagraphPortion portion) => portion.Text == Environment.NewLine;

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

    private static float GetBaselineOffset(SKFont font)
    {
        var ascent = font.Metrics.Ascent;
        return ascent >= 0 ? 0 : -ascent;
    }

    private static decimal ClampToZero(decimal value)
    {
        return value < 0 ? 0 : value;
    }

    private static IEnumerable<string> SplitToFittingParts(string token, SKFont font, float maxWidth)
    {
        var offset = 0;
        while (offset < token.Length)
        {
            var partLength = GetFittingPartLength(token, offset, font, maxWidth);
            yield return token.Substring(offset, partLength);
            offset += partLength;
        }
    }

    private static int GetFittingPartLength(string token, int offset, SKFont font, float maxWidth)
    {
        var remaining = token.Length - offset;
        if (remaining <= 0)
        {
            return 0;
        }

        if (maxWidth <= 0)
        {
            return remaining;
        }

        var low = 1;
        var high = remaining;
        var best = 0;
        var tokenSpan = token.AsSpan();

        while (low <= high)
        {
            var mid = low + ((high - low) / 2);
            var candidate = tokenSpan.Slice(offset, mid);
            var width = font.MeasureText(candidate);

            if (width <= maxWidth)
            {
                best = mid;
                low = mid + 1;
            }
            else
            {
                high = mid - 1;
            }
        }

        return best == 0 ? 1 : best;
    }
    
    private static void RenderLine(SKCanvas canvas, TextLine line, float startX, float lineTop)
    {
        var baselineY = lineTop + line.BaselineOffset;
        var currentX = startX;

        foreach (var run in line.Runs)
        {
            using var font = CreateFont(run.Font);
            using var paint = CreatePaint(run.Font);

            canvas.DrawText(run.Text, currentX, baselineY, SKTextAlign.Left, font, paint);
            currentX += run.Width;
        }
    }

    private static string[] SplitToLineSegments(string text)
    {
        return text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
    }

    private float GetAvailableWidth(decimal parentShapeWidth)
    {
        var width = ClampToZero(parentShapeWidth - this.LeftMargin - this.RightMargin);
        return (float)new Points(width).AsPixels();
    }

    private float GetAvailableHeight(decimal parentShapeHeight)
    {
        var height = ClampToZero(parentShapeHeight - this.TopMargin - this.BottomMargin);
        return (float)new Points(height).AsPixels();
    }

    private IReadOnlyList<TextLine> LayoutLines(float availableWidth, bool wrap)
    {
        var lines = new List<TextLine>();

        foreach (var paragraph in this.Paragraphs)
        {
            LayoutParagraph(paragraph, availableWidth, wrap, lines);
        }

        return lines;
    }

    private void LayoutParagraph(IParagraph paragraph, float availableWidth, bool wrap, ICollection<TextLine> buffer)
    {
        var paraLeftMargin = (float)new Points(paragraph.LeftMargin).AsPixels();
        var firstLineIndent = (float)new Points(paragraph.FirstLineIndent).AsPixels();
        var firstLineLeftMargin = paraLeftMargin + firstLineIndent;
        var line = new LineBuilder(paragraph.HorizontalAlignment, firstLineLeftMargin);
        if (paragraph.Bullet.Type == BulletType.Character && paragraph.Bullet.Character != null)
        {
            var font = paragraph.Portions.FirstOrDefault()?.Font;
            if (font != null)
            {
                using var skFont = CreateFont(font);
                var bulletChar = paragraph.Bullet.Character;
                var bulletCharWidth = skFont.MeasureText(bulletChar);
                var hangingIndentWidth = firstLineIndent < 0 ? -firstLineIndent : 0f;
                var bulletPortionWidth = Math.Max(bulletCharWidth, hangingIndentWidth);
                var bulletPortion = new PixelTextPortion(bulletChar, font, bulletPortionWidth);

                line.Add(bulletPortion, skFont.Spacing, GetBaselineOffset(skFont));
            }
        }

        foreach (var portion in paragraph.Portions)
        {
            if (IsLineBreak(portion))
            {
                buffer.Add(line.Build(defaultLineHeight, defaultBaselineOffset));
                line = new LineBuilder(paragraph.HorizontalAlignment, paraLeftMargin);
                continue;
            }

            line = LayoutTextPortion(portion, line, paragraph.HorizontalAlignment, availableWidth, paraLeftMargin, wrap, buffer);
        }

        buffer.Add(line.Build(defaultLineHeight, defaultBaselineOffset));
    }

    private LineBuilder LayoutTextPortion(
        IParagraphPortion portion,
        LineBuilder currentLine,
        TextHorizontalAlignment paragraphAlignment,
        float totalAvailableWidth,
        float baseParaLeftMargin,
        bool wrap,
        ICollection<TextLine> buffer)
    {
        using var font = CreateFont(portion.Font);
        var baselineOffset = GetBaselineOffset(font);
        var paragraphLayout = new ParagraphLayout(paragraphAlignment, totalAvailableWidth, baseParaLeftMargin, wrap, buffer);
        
        // Open XML text runs can contain hard line breaks as '\r'/'\n' inside <a:t>.
        // PowerPoint treats them as explicit new lines; Skia would otherwise render them as tofu.
        var segments = SplitToLineSegments(portion.Text);
        for (var i = 0; i < segments.Length; i++)
        {
            foreach (var token in SplitToTokens(segments[i]))
            {
                currentLine = LayoutToken(
                    token,
                    portion.Font,
                    font,
                    baselineOffset,
                    paragraphLayout,
                    currentLine);
            }
            
            if (i < segments.Length - 1)
            {
                paragraphLayout.Buffer.Add(currentLine.Build(defaultLineHeight, defaultBaselineOffset));
                currentLine = new LineBuilder(paragraphLayout.ParagraphAlignment, baseParaLeftMargin);
            }
        }

        return currentLine;
    }

    private LineBuilder LayoutToken(
        string token,
        ITextPortionFont? font,
        SKFont skFont,
        float baselineOffset,
        ParagraphLayout paragraphLayout,
        LineBuilder currentLine)
    {
        var tokenWidth = skFont.MeasureText(token);

        if (!paragraphLayout.Wrap || paragraphLayout.TotalAvailableWidth <= 0)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), skFont.Spacing, baselineOffset);
            return currentLine;
        }

        if (IsWhitespace(token))
        {
            return LayoutWhitespaceToken(
                token,
                font,
                tokenWidth,
                skFont.Spacing,
                baselineOffset,
                paragraphLayout,
                currentLine);
        }

        return LayoutNonWhitespaceToken(
            token,
            font,
            skFont,
            tokenWidth,
            baselineOffset,
            paragraphLayout,
            currentLine);
    }

    private LineBuilder LayoutWhitespaceToken(
        string token,
        ITextPortionFont? font,
        float tokenWidth,
        float spacing,
        float baselineOffset,
        ParagraphLayout paragraphLayout,
        LineBuilder currentLine)
    {
        if (currentLine.ParaLeftMargin + currentLine.Width + tokenWidth <= paragraphLayout.TotalAvailableWidth)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), spacing, baselineOffset);
            return currentLine;
        }

        paragraphLayout.Buffer.Add(currentLine.Build(defaultLineHeight, defaultBaselineOffset));
        return new LineBuilder(paragraphLayout.ParagraphAlignment, paragraphLayout.BaseParaLeftMargin); // Drop whitespace at wrap boundary.
    }

    private LineBuilder LayoutNonWhitespaceToken(
        string token,
        ITextPortionFont? font,
        SKFont skFont,
        float tokenWidth,
        float baselineOffset,
        ParagraphLayout paragraphLayout,
        LineBuilder currentLine)
    {
        if (currentLine.ParaLeftMargin + currentLine.Width + tokenWidth <= paragraphLayout.TotalAvailableWidth)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), skFont.Spacing, baselineOffset);
            return currentLine;
        }

        if (currentLine.HasRuns)
        {
            paragraphLayout.Buffer.Add(currentLine.Build(defaultLineHeight, defaultBaselineOffset));
            currentLine = new LineBuilder(paragraphLayout.ParagraphAlignment, paragraphLayout.BaseParaLeftMargin);
        }

        if (currentLine.ParaLeftMargin + tokenWidth <= paragraphLayout.TotalAvailableWidth)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), skFont.Spacing, baselineOffset);
            return currentLine;
        }

        return LayoutSplitToken(
            token,
            font,
            skFont,
            baselineOffset,
            paragraphLayout,
            currentLine);
    }

    private LineBuilder LayoutSplitToken(
        string token,
        ITextPortionFont? font,
        SKFont skFont,
        float baselineOffset,
        ParagraphLayout paragraphLayout,
        LineBuilder currentLine)
    {
        var remainingToken = token;
        while (remainingToken.Length > 0)
        {
            var availableWidthForLine = paragraphLayout.TotalAvailableWidth - (currentLine.ParaLeftMargin + currentLine.Width);
            if (availableWidthForLine <= 0 && currentLine.HasRuns)
            {
                paragraphLayout.Buffer.Add(currentLine.Build(defaultLineHeight, defaultBaselineOffset));
                currentLine = new LineBuilder(paragraphLayout.ParagraphAlignment, paragraphLayout.BaseParaLeftMargin);
                availableWidthForLine = paragraphLayout.TotalAvailableWidth - currentLine.ParaLeftMargin;
            }

            if (availableWidthForLine <= 0)
            {
                // No available width even on an empty line: force the remaining token into this line
                var forcedPart = remainingToken;
                var forcedPartWidth = skFont.MeasureText(forcedPart);
                currentLine.Add(new PixelTextPortion(forcedPart, font, forcedPartWidth), skFont.Spacing, baselineOffset);
                break;
            }

            var partLength = GetFittingPartLength(remainingToken, 0, skFont, availableWidthForLine);
            var part = remainingToken[..partLength];
            var partWidth = skFont.MeasureText(part);

            currentLine.Add(new PixelTextPortion(part, font, partWidth), skFont.Spacing, baselineOffset);
            remainingToken = remainingToken[partLength..];
            
            if (remainingToken.Length > 0)
            {
                paragraphLayout.Buffer.Add(currentLine.Build(defaultLineHeight, defaultBaselineOffset));
                currentLine = new LineBuilder(paragraphLayout.ParagraphAlignment, paragraphLayout.BaseParaLeftMargin);
            }
        }

        return currentLine;
    }

    private readonly struct ParagraphLayout
    {
        internal ParagraphLayout(
            TextHorizontalAlignment paragraphAlignment,
            float totalAvailableWidth,
            float baseParaLeftMargin,
            bool wrap,
            ICollection<TextLine> buffer)
        {
            ParagraphAlignment = paragraphAlignment;
            TotalAvailableWidth = totalAvailableWidth;
            BaseParaLeftMargin = baseParaLeftMargin;
            Wrap = wrap;
            Buffer = buffer;
        }

        internal TextHorizontalAlignment ParagraphAlignment { get; }

        internal float TotalAvailableWidth { get; }

        internal float BaseParaLeftMargin { get; }

        internal bool Wrap { get; }

        internal ICollection<TextLine> Buffer { get; }
    }

    private sealed class TextLine
    {
        internal TextLine(
            PixelTextPortion[] runs,
            TextHorizontalAlignment horizontalAlignment,
            float paraLeftMargin,
            float width,
            float height,
            float baselineOffset)
        {
            Runs = runs;
            HorizontalAlignment = horizontalAlignment;
            ParaLeftMargin = paraLeftMargin;
            Width = width;
            Height = height;
            BaselineOffset = baselineOffset;
        }

        internal PixelTextPortion[] Runs { get; }

        internal TextHorizontalAlignment HorizontalAlignment { get; }

        internal float ParaLeftMargin { get; }

        internal float Width { get; }

        internal float Height { get; }

        internal float BaselineOffset { get; }
    }

    private sealed class LineBuilder(TextHorizontalAlignment paragraphAlignment, float paraLeftMargin)
    {
        private readonly List<PixelTextPortion> runs = [];

        internal float ParaLeftMargin => paraLeftMargin;

        internal float Width { get; private set; }

        internal bool HasRuns => runs.Count > 0;

        private float Height { get; set; }

        private float BaselineOffset { get; set; }

        internal void Add(PixelTextPortion portion, float spacing, float baselineOffset)
        {
            runs.Add(portion);
            Width += portion.Width;
            Height = Math.Max(Height, spacing);
            BaselineOffset = Math.Max(BaselineOffset, baselineOffset);
        }

        internal TextLine Build(float defaultHeight, float defaultBaselineOffset)
        {
            TrimTrailingWhitespace();

            var height = Height <= 0 ? defaultHeight : Height;
            var baselineOffset = BaselineOffset <= 0 ? defaultBaselineOffset : BaselineOffset;
            return new TextLine([.. runs], paragraphAlignment, paraLeftMargin, Width, height, baselineOffset);
        }

        private void TrimTrailingWhitespace()
        {
            while (runs.Count > 0 && runs[^1].IsWhitespace)
            {
                var removingRun = runs[^1];
                Width -= removingRun.Width;
                runs.RemoveAt(runs.Count - 1);
            }
        }
    }
}