using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Units;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Represents a text layout.
/// </summary>
/// <param name="paragraphs">The paragraphs to be laid out into lines.</param>
/// <param name="availableWidth">The maximum available width for each rendered line.</param>
/// <param name="wrap">True to wrap text across multiple lines when it exceeds the available width; otherwise, false.</param>
internal struct TextLayout(IReadOnlyList<IParagraph> paragraphs, float availableWidth, bool wrap)
{
    private float defaultLineHeight;
    private float defaultBaselineOffset;

    internal void Render(SKCanvas canvas, float x, float y, float availableHeight, TextVerticalAlignment verticalAlignment)
    {
        var drawingFont = new DrawingFont(null);
        using var font = drawingFont.AsSkFont();
        this.defaultLineHeight = font.Spacing;
        this.defaultBaselineOffset = DrawingFont.BaselineOffset(font);

        var lines = this.LayoutLines();
        new TextLines(lines, availableWidth).Render(canvas, x, y, availableHeight, verticalAlignment);
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

    private static bool IsLineBreak(IParagraphPortion portion) => portion.Text == Environment.NewLine;

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

    private static string[] SplitToLineSegments(string text)
    {
        return text.Split(["\r\n", "\n", "\r"], StringSplitOptions.None);
    }

    private static string NormalizeTextForRendering(string text)
    {
        // PowerPoint can contain U+2011 (non-breaking hyphen) which many fonts don't expose as a glyph.
        // PowerPoint renders it like a regular hyphen, so do the same to avoid tofu in Skia output.
        return text.IndexOf('\u2011') >= 0
            ? text.Replace('\u2011', '-')
            : text;
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
            return this.LayoutWhitespaceToken(
                token,
                font,
                tokenWidth,
                skFont.Spacing,
                baselineOffset,
                paragraphLayout,
                currentLine);
        }

        return this.LayoutNonWhitespaceToken(
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

        paragraphLayout.Buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
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
            paragraphLayout.Buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
            currentLine = new LineBuilder(paragraphLayout.ParagraphAlignment, paragraphLayout.BaseParaLeftMargin);
        }

        if (currentLine.ParaLeftMargin + tokenWidth <= paragraphLayout.TotalAvailableWidth)
        {
            currentLine.Add(new PixelTextPortion(token, font, tokenWidth), skFont.Spacing, baselineOffset);
            return currentLine;
        }

        return this.LayoutSplitToken(
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
                paragraphLayout.Buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
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
                paragraphLayout.Buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
                currentLine = new LineBuilder(paragraphLayout.ParagraphAlignment, paragraphLayout.BaseParaLeftMargin);
            }
        }

        return currentLine;
    }

    private List<TextLine> LayoutLines()
    {
        var lines = new List<TextLine>();

        foreach (var paragraph in paragraphs)
        {
            this.LayoutParagraph(paragraph, lines);
        }

        return lines;
    }

    private void LayoutParagraph(IParagraph paragraph, ICollection<TextLine> textLines)
    {
        var paraLeftMargin = (float)new Points(paragraph.LeftMargin).AsPixels();
        var firstLineIndent = (float)new Points(paragraph.FirstLineIndent).AsPixels();
        var firstLineLeftMargin = paraLeftMargin + firstLineIndent;
        var line = new LineBuilder(paragraph.HorizontalAlignment, firstLineLeftMargin);
        if (paragraph.Bullet is { Type: BulletType.Character, Character: not null })
        {
            var font = paragraph.Portions.FirstOrDefault()?.Font;
            if (font != null)
            {
                var drawingFont = new DrawingFont(font);
                using var skFont = drawingFont.AsSkFont();
                var bulletChar = paragraph.Bullet.Character;
                var bulletCharWidth = skFont.MeasureText(bulletChar);
                var hangingIndentWidth = firstLineIndent < 0 ? -firstLineIndent : 0f;
                var bulletPortionWidth = Math.Max(bulletCharWidth, hangingIndentWidth);
                var bulletPortion = new PixelTextPortion(bulletChar, font, bulletPortionWidth);

                line.Add(bulletPortion, skFont.Spacing, DrawingFont.BaselineOffset(skFont));
            }
        }

        foreach (var portion in paragraph.Portions)
        {
            if (IsLineBreak(portion))
            {
                textLines.Add(line.Build(this.defaultLineHeight, this.defaultBaselineOffset));
                line = new LineBuilder(paragraph.HorizontalAlignment, paraLeftMargin);
                continue;
            }

            line = this.LayoutTextPortion(portion, line, paragraph.HorizontalAlignment, paraLeftMargin, textLines);
        }

        textLines.Add(line.Build(this.defaultLineHeight, this.defaultBaselineOffset));
    }

    private LineBuilder LayoutTextPortion(
        IParagraphPortion portion,
        LineBuilder currentLine,
        TextHorizontalAlignment paragraphAlignment,
        float baseParaLeftMargin,
        ICollection<TextLine> buffer)
    {
        var drawingFont = new DrawingFont(portion.Font);
        using var font = drawingFont.AsSkFont();
        var baselineOffset = DrawingFont.BaselineOffset(font);
        var paragraphLayout = new ParagraphLayout(paragraphAlignment, availableWidth, baseParaLeftMargin, wrap, buffer);

        // Open XML text runs can contain hard line breaks as '\r'/'\n' inside <a:t>.
        // PowerPoint treats them as explicit new lines; Skia would otherwise render them as tofu.
        var normalizedText = NormalizeTextForRendering(portion.Text);
        var segments = SplitToLineSegments(normalizedText);
        for (var i = 0; i < segments.Length; i++)
        {
            foreach (var token in SplitToTokens(segments[i]))
            {
                currentLine = this.LayoutToken(
                    token,
                    portion.Font,
                    font,
                    baselineOffset,
                    paragraphLayout,
                    currentLine);
            }

            if (i < segments.Length - 1)
            {
                paragraphLayout.Buffer.Add(currentLine.Build(this.defaultLineHeight, this.defaultBaselineOffset));
                currentLine = new LineBuilder(paragraphLayout.ParagraphAlignment, baseParaLeftMargin);
            }
        }

        return currentLine;
    }

    private readonly struct ParagraphLayout(
        TextHorizontalAlignment paragraphAlignment,
        float totalAvailableWidth,
        float baseParaLeftMargin,
        bool wrap,
        ICollection<TextLine> buffer)
    {
        internal TextHorizontalAlignment ParagraphAlignment { get; } = paragraphAlignment;

        internal float TotalAvailableWidth { get; } = totalAvailableWidth;

        internal float BaseParaLeftMargin { get; } = baseParaLeftMargin;

        internal bool Wrap { get; } = wrap;

        internal ICollection<TextLine> Buffer { get; } = buffer;
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