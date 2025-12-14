using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents text drawing.
/// </summary>
internal sealed class TextDrawing
{
    private const decimal DefaultFontSize = 12m;
    private readonly float defaultLineHeight;
    private readonly float defaultBaselineOffset;

    internal TextDrawing()
    {
        using var font = CreateFont(null);
        this.defaultLineHeight = font.Spacing;
        this.defaultBaselineOffset = GetBaselineOffset(font);
    }

    internal void Render(SKCanvas canvas, IShape shape)
    {
        var textBox = shape.TextBox;
        if (textBox is null || string.IsNullOrWhiteSpace(textBox.Text))
        {
            return;
        }

        var originX = (float)new Points(shape.X + textBox.LeftMargin).AsPixels();
        var originY = (float)new Points(shape.Y + textBox.TopMargin).AsPixels();
        var availableWidth = GetAvailableWidth(shape, textBox);
        var availableHeight = GetAvailableHeight(shape, textBox);

        var wrap = textBox.TextWrapped && availableWidth > 0;
        var lines = this.LayoutLines(textBox, availableWidth, wrap);
        var textBlockHeight = lines.Sum(l => l.Height);

        var verticalOffset = GetVerticalOffset(textBox.VerticalAlignment, availableHeight, textBlockHeight);
        var lineTop = originY + verticalOffset;

        foreach (var line in lines)
        {
            var horizontalOffset = GetHorizontalOffset(line.HorizontalAlignment, availableWidth, line.Width);
            var startX = originX + horizontalOffset;

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

    private static float GetAvailableWidth(IShape shape, ITextBox textBox)
    {
        var width = ClampToZero(shape.Width - textBox.LeftMargin - textBox.RightMargin);
        return (float)new Points(width).AsPixels();
    }

    private static float GetAvailableHeight(IShape shape, ITextBox textBox)
    {
        var height = ClampToZero(shape.Height - textBox.TopMargin - textBox.BottomMargin);
        return (float)new Points(height).AsPixels();
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

    private IReadOnlyList<TextLine> LayoutLines(ITextBox textBox, float availableWidth, bool wrap)
    {
        var lines = new List<TextLine>();

        foreach (var paragraph in textBox.Paragraphs)
        {
            this.LayoutParagraph(paragraph, availableWidth, wrap, lines);
        }

        return lines;
    }

    private void LayoutParagraph(IParagraph paragraph, float availableWidth, bool wrap, ICollection<TextLine> buffer)
    {
        var line = new LineBuilder(paragraph.HorizontalAlignment);

        foreach (var portion in paragraph.Portions)
        {
            if (IsLineBreak(portion))
            {
                buffer.Add(line.Build(this.defaultLineHeight, this.defaultBaselineOffset));
                line = new LineBuilder(paragraph.HorizontalAlignment);
                continue;
            }

            line = this.LayoutTextPortion(portion, line, paragraph.HorizontalAlignment, availableWidth, wrap, buffer);
        }

        buffer.Add(line.Build(this.defaultLineHeight, this.defaultBaselineOffset));
    }

    private LineBuilder LayoutTextPortion(
        IParagraphPortion portion,
        LineBuilder currentLine,
        TextHorizontalAlignment paragraphAlignment,
        float availableWidth,
        bool wrap,
        ICollection<TextLine> buffer)
    {
        using var font = CreateFont(portion.Font);
        var baselineOffset = GetBaselineOffset(font);

        foreach (var token in SplitToTokens(portion.Text))
        {
            currentLine = this.LayoutToken(
                token,
                portion.Font,
                font,
                baselineOffset,
                paragraphAlignment,
                availableWidth,
                wrap,
                currentLine,
                buffer);
        }

        return currentLine;
    }

    private LineBuilder LayoutToken(
        string token,
        ITextPortionFont? font,
        SKFont skFont,
        float baselineOffset,
        TextHorizontalAlignment paragraphAlignment,
        float availableWidth,
        bool wrap,
        LineBuilder currentLine,
        ICollection<TextLine> buffer)
    {
        var tokenWidth = skFont.MeasureText(token);

        if (!wrap || availableWidth <= 0)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), skFont.Spacing, baselineOffset);
            return currentLine;
        }

        if (IsWhitespace(token))
        {
            return this.LayoutWhitespaceToken(
                token,
                font,
                tokenWidth,
                skFont.Spacing,
                baselineOffset,
                paragraphAlignment,
                availableWidth,
                currentLine,
                buffer);
        }

        return this.LayoutNonWhitespaceToken(
            token,
            font,
            skFont,
            tokenWidth,
            baselineOffset,
            paragraphAlignment,
            availableWidth,
            currentLine,
            buffer);
    }

    private LineBuilder LayoutWhitespaceToken(
        string token,
        ITextPortionFont? font,
        float tokenWidth,
        float spacing,
        float baselineOffset,
        TextHorizontalAlignment paragraphAlignment,
        float availableWidth,
        LineBuilder currentLine,
        ICollection<TextLine> buffer)
    {
        if (currentLine.Width + tokenWidth <= availableWidth)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), spacing, baselineOffset);
            return currentLine;
        }

        buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
        return new LineBuilder(paragraphAlignment); // Drop whitespace at wrap boundary.
    }

    private LineBuilder LayoutNonWhitespaceToken(
        string token,
        ITextPortionFont? font,
        SKFont skFont,
        float tokenWidth,
        float baselineOffset,
        TextHorizontalAlignment paragraphAlignment,
        float availableWidth,
        LineBuilder currentLine,
        ICollection<TextLine> buffer)
    {
        if (currentLine.Width + tokenWidth <= availableWidth)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), skFont.Spacing, baselineOffset);
            return currentLine;
        }

        if (currentLine.HasRuns)
        {
            buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
            currentLine = new LineBuilder(paragraphAlignment);
        }

        if (tokenWidth <= availableWidth)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), skFont.Spacing, baselineOffset);
            return currentLine;
        }

        return this.LayoutSplitToken(
            token,
            font,
            skFont,
            baselineOffset,
            paragraphAlignment,
            availableWidth,
            currentLine,
            buffer);
    }

    private LineBuilder LayoutSplitToken(
        string token,
        ITextPortionFont? font,
        SKFont skFont,
        float baselineOffset,
        TextHorizontalAlignment paragraphAlignment,
        float availableWidth,
        LineBuilder currentLine,
        ICollection<TextLine> buffer)
    {
        foreach (var part in SplitToFittingParts(token, skFont, availableWidth))
        {
            var partWidth = skFont.MeasureText(part);

            if (currentLine.HasRuns && currentLine.Width + partWidth > availableWidth)
            {
                buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
                currentLine = new LineBuilder(paragraphAlignment);
            }

            currentLine.Add(new PixelTextPortion(part, font, partWidth), skFont.Spacing, baselineOffset);
        }

        return currentLine;
    }

    private sealed class TextLine
    {
        internal TextLine(
            PixelTextPortion[] runs,
            TextHorizontalAlignment horizontalAlignment,
            float width,
            float height,
            float baselineOffset)
        {
            this.Runs = runs;
            this.HorizontalAlignment = horizontalAlignment;
            this.Width = width;
            this.Height = height;
            this.BaselineOffset = baselineOffset;
        }

        internal PixelTextPortion[] Runs { get; }

        internal TextHorizontalAlignment HorizontalAlignment { get; }

        internal float Width { get; }

        internal float Height { get; }

        internal float BaselineOffset { get; }
    }

    private sealed class LineBuilder
    {
        private readonly List<PixelTextPortion> runs;
        private readonly TextHorizontalAlignment paragraphAlignment;


        internal LineBuilder(TextHorizontalAlignment paragraphAlignment)
        {
            this.paragraphAlignment = paragraphAlignment;
            this.runs = [];
        }

        internal float Width { get; private set; }

        internal bool HasRuns => this.runs.Count > 0;

        private float Height { get; set; }

        private float BaselineOffset { get; set; }

        internal void Add(PixelTextPortion portion, float spacing, float baselineOffset)
        {
            this.runs.Add(portion);
            this.Width += portion.Width;
            this.Height = Math.Max(this.Height, spacing);
            this.BaselineOffset = Math.Max(this.BaselineOffset, baselineOffset);
        }

        internal TextLine Build(float defaultHeight, float defaultBaselineOffset)
        {
            this.TrimTrailingWhitespace();

            var height = this.Height <= 0 ? defaultHeight : this.Height;
            var baselineOffset = this.BaselineOffset <= 0 ? defaultBaselineOffset : this.BaselineOffset;
            return new TextLine([.. this.runs], this.paragraphAlignment, this.Width, height, baselineOffset);
        }

        private void TrimTrailingWhitespace()
        {
            while (this.runs.Count > 0 && this.runs[^1].IsWhitespace)
            {
                var removingRun = this.runs[^1];
                this.Width -= removingRun.Width;
                this.runs.RemoveAt(this.runs.Count - 1);
            }
        }
    }
}