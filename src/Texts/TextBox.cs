using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Texts;

internal sealed class TextBox : ITextBox
{
    private readonly TextBoxMargins margins;
    private readonly OpenXmlElement textBody;
    private readonly ShapeSize shapeSize;
    private TextVerticalAlignment? vAlignment;

    internal TextBox(TextBoxMargins margins, OpenXmlElement textBody)
    {
        this.margins = margins;
        this.textBody = textBody;
        this.shapeSize = new ShapeSize(textBody.Parent!);
    }

    public IParagraphCollection Paragraphs => new ParagraphCollection(this.textBody);

    public string Text
    {
        get
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append(this.Paragraphs[0].Text);

            var paragraphsCount = this.Paragraphs.Count;
            var index = 1; // we've already added the text of first paragraph
            while (index < paragraphsCount)
            {
                stringBuilder.AppendLine();
                stringBuilder.Append(this.Paragraphs[index].Text);

                index++;
            }

            return stringBuilder.ToString();
        }
    }

    public AutofitType AutofitType
    {
        get
        {
            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.GetFirstChild<A.NormalAutoFit>() != null)
            {
                return AutofitType.Shrink;
            }

            if (aBodyPr.GetFirstChild<A.ShapeAutoFit>() != null)
            {
                return AutofitType.Resize;
            }

            return AutofitType.None;
        }

        set
        {
            var currentType = this.AutofitType;
            if (currentType == value)
            {
                return;
            }

            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>() !;

            RemoveExistingAutofitElements(aBodyPr);

            switch (value)
            {
                case AutofitType.None:
                    aBodyPr.Append(new A.NoAutoFit());
                    break;
                case AutofitType.Shrink:
                    aBodyPr.Append(new A.NormalAutoFit());
                    break;
                case AutofitType.Resize:
                    aBodyPr.Append(new A.ShapeAutoFit());
                    this.ResizeParentShapeOnDemand();
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }

    public decimal LeftMargin
    {
        get => this.margins.Left;
        set => this.margins.Left = value;
    }

    public decimal RightMargin
    {
        get => this.margins.Right;
        set => this.margins.Right = value;
    }

    public decimal TopMargin
    {
        get => this.margins.Top;
        set
        {
            if (this.margins != null)
            {
                this.margins.Top = value;
            }
        }
    }

    public decimal BottomMargin
    {
        get => this.margins.Bottom;
        set => this.margins.Bottom = value;
    }

    public bool TextWrapped
    {
        get
        {
            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>() !;
            var wrap = aBodyPr.GetAttributes().FirstOrDefault(a => a.LocalName == "wrap");

            if (wrap.Value == "none")
            {
                return false;
            }

            return true;
        }
    }

    public string SDKXPath => new XmlPath(this.textBody).XPath;

    public TextVerticalAlignment VerticalAlignment
    {
        get
        {
            if (this.vAlignment.HasValue)
            {
                return this.vAlignment.Value;
            }

            var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>();

            if (aBodyPr!.Anchor?.Value == A.TextAnchoringTypeValues.Center)
            {
                this.vAlignment = TextVerticalAlignment.Middle;
            }
            else if (aBodyPr.Anchor?.Value == A.TextAnchoringTypeValues.Bottom)
            {
                this.vAlignment = TextVerticalAlignment.Bottom;
            }
            else
            {
                this.vAlignment = TextVerticalAlignment.Top;
            }

            return this.vAlignment.Value;
        }

        set => this.SetVerticalAlignment(value);
    }

    public void SetMarkdownText(string text)
    {
        var lines = Regex.Split(text, "\r\n|\r|\n", RegexOptions.None, TimeSpan.FromMilliseconds(1000));
        if (IsList(lines))
        {
            this.RenderList(lines);
        }
        else
        {
            this.RenderRegularText(text);
        }

        this.ResizeParentShapeOnDemand();
    }

    public void SetText(string text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var paragraphs = this.Paragraphs.ToList();
        var firstParagraph = paragraphs.FirstOrDefault();

        // Store LatinName from first portion if available
        string? latinNameToPreserve = GetLatinNameToPreserve(firstParagraph);

        // Store font color hex from first portion if available
        string? colorHexToPreserve = GetFontColorHexToPreserve(firstParagraph);

        // Clear existing content and ensure we have a first paragraph
        firstParagraph = this.PrepareTextContainer(firstParagraph, paragraphs);

        // Add new text with preserved font
        var paragraphLines = text.Split([Environment.NewLine], StringSplitOptions.None);
        this.AddTextToParagraphs(paragraphLines, firstParagraph, latinNameToPreserve);
        if (colorHexToPreserve != null)
        {
            for (int i = 0; i < paragraphLines.Length; i++)
            {
                var portion = this.Paragraphs[i].Portions.Last();
                portion.Font!.Color.Set(colorHexToPreserve);
            }
        }

        this.ApplyTextFormatting(text);
    }

    internal void ResizeParentShapeOnDemand()
    {
        if (this.AutofitType != AutofitType.Resize)
        {
            return;
        }

        var shapeWidthCapacity = this.shapeSize.Width - this.LeftMargin - this.RightMargin;
        var shapeHeightCapacity = this.shapeSize.Height - this.TopMargin - this.BottomMargin;

        decimal textHeightPx = 0;
        foreach (var paragraph in this.Paragraphs)
        {
            var paragraphPortion = paragraph.Portions.OfType<TextParagraphPortion>();
            if (!paragraphPortion.Any())
            {
                continue;
            }

            var popularPortion = paragraphPortion.GroupBy(p => p.Font.Size)
                .OrderByDescending(x => x.Count())
                .First().First();
            var scFont = popularPortion.Font;

            var paragraphText = paragraph.Text.ToUpper();
            var paragraphTextWidth = new Text(paragraphText, scFont).Width;
            var paragraphTextHeight = scFont.Size;
            var requiredRowsCount = paragraphTextWidth / shapeWidthCapacity;
            var intRequiredRowsCount = (int)requiredRowsCount;
            var fractionalPart = requiredRowsCount - intRequiredRowsCount;
            if (fractionalPart > 0)
            {
                intRequiredRowsCount++;
            }

            textHeightPx += intRequiredRowsCount * (int)paragraphTextHeight;
        }

        this.UpdateShapeHeight(textHeightPx, shapeHeightCapacity);
        if (!this.TextWrapped)
        {
            this.UpdateShapeWidth();
        }
    }

    private static bool IsList(string[] lines) =>
        lines.Any(l => l.TrimStart().StartsWith("- ", StringComparison.CurrentCulture));

    private static void RemoveExistingAutofitElements(A.BodyProperties bodyProperties)
    {
        bodyProperties.GetFirstChild<A.NoAutoFit>()?.Remove();
        bodyProperties.GetFirstChild<A.NormalAutoFit>()?.Remove();
        bodyProperties.GetFirstChild<A.ShapeAutoFit>()?.Remove();
    }

    private static void ApplyLatinNameIfNeeded(IParagraphPortion portion, string? latinNameToPreserve)
    {
        if (latinNameToPreserve != null && portion.Font != null)
        {
            portion.Font.LatinName = latinNameToPreserve;
        }
    }

    private static string? GetLatinNameToPreserve(IParagraph? firstParagraph)
    {
        var firstPortion = firstParagraph?.Portions.FirstOrDefault();
        return firstPortion?.Font?.LatinName;
    }

    private static string? GetFontColorHexToPreserve(IParagraph? firstParagraph)
    {
        var firstPortion = firstParagraph?.Portions.FirstOrDefault();
        return firstPortion?.Font?.Color.Hex;
    }

    private IParagraph PrepareTextContainer(
        IParagraph? firstParagraph,
        System.Collections.Generic.List<IParagraph> paragraphs)
    {
        if (firstParagraph == null)
        {
            this.Paragraphs.Add();
            return this.Paragraphs.First();
        }

        // Remove all paragraphs except the first one
        foreach (var paragraph in paragraphs.Skip(1))
        {
            paragraph.Remove();
        }

        // Clear the first paragraph
        foreach (var portion in firstParagraph.Portions.ToList())
        {
            portion.Remove();
        }

        return firstParagraph;
    }

    private void AddTextToParagraphs(string[] paragraphLines, IParagraph firstParagraph, string? latinNameToPreserve)
    {
        if (paragraphLines.Length <= 0)
        {
            return;
        }

        // Add first line to the first paragraph
        firstParagraph.Portions.AddText(paragraphLines[0]);
        ApplyLatinNameIfNeeded(firstParagraph.Portions.Last(), latinNameToPreserve);

        // Add remaining lines as new paragraphs
        for (int i = 1; i < paragraphLines.Length; i++)
        {
            this.Paragraphs.Add();
            this.Paragraphs[i].Portions.AddText(paragraphLines[i]);
            ApplyLatinNameIfNeeded(this.Paragraphs[i].Portions.Last(), latinNameToPreserve);
        }
    }

    private void ApplyTextFormatting(string text)
    {
        if (this.AutofitType == AutofitType.Shrink)
        {
            this.ShrinkFont(text);
        }

        this.ResizeParentShapeOnDemand();
    }

    private void SetVerticalAlignment(TextVerticalAlignment alignmentValue)
    {
        var aTextAlignmentTypeValue = alignmentValue switch
        {
            TextVerticalAlignment.Top => A.TextAnchoringTypeValues.Top,
            TextVerticalAlignment.Middle => A.TextAnchoringTypeValues.Center,
            TextVerticalAlignment.Bottom => A.TextAnchoringTypeValues.Bottom,
            _ => throw new ArgumentOutOfRangeException(nameof(alignmentValue))
        };

        var aBodyPr = this.textBody.GetFirstChild<A.BodyProperties>();

        if (aBodyPr is not null)
        {
            aBodyPr.Anchor = aTextAlignmentTypeValue;
        }

        this.vAlignment = alignmentValue;
    }

    private void ShrinkFont(string newText)
    {
        var firstParagraph = this.Paragraphs.First();
        var popularFont = firstParagraph.Portions.GroupBy(paraPortion => paraPortion.Font!.Size)
            .OrderByDescending(x => x.Count())
            .First().First().Font!;
        var text = new Text(newText, popularFont);
        text.Fit(this.shapeSize.Width, this.shapeSize.Height);
        firstParagraph.SetFontSize((int)text.FontSize);
    }

    private void UpdateShapeWidth()
    {
        var longerText = this.Paragraphs
            .Select(x => new { x.Text, x.Text.Length })
            .OrderByDescending(x => x.Length)
            .First().Text;

        var baseParagraph = this.Paragraphs.First();
        var popularPortion = baseParagraph.Portions.OfType<TextParagraphPortion>().GroupBy(p => p.Font.Size)
            .OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var textWidth = new Text(longerText, font).Width;
        var leftMargin = this.LeftMargin;
        var rightMargin = this.RightMargin;
        var newWidth =
            (int)(textWidth *
                  1.4M) // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi. 96/72 = 1.4 
            + leftMargin + rightMargin;
        this.shapeSize.Width = newWidth;
    }

    private void UpdateShapeHeight(decimal textHeight, decimal shapeHeightPtCapacity)
    {
        var parentShape = this.textBody.Parent!;
        var requiredHeightPt = textHeight + this.TopMargin + this.BottomMargin;
        var newHeight = requiredHeightPt + this.TopMargin + this.BottomMargin + this.TopMargin + this.BottomMargin;
        this.shapeSize.Height = newHeight;

        // Raise the shape up by the amount, which is half of the increased offset, like PowerPoint does it
        var position = new Position(parentShape);
        var yOffset = (requiredHeightPt - shapeHeightPtCapacity) / 2;
        position.Y -= yOffset;
    }

    private void RenderList(string[] lines)
    {
        var paragraphs = this.Paragraphs.ToList();
        var firstPara = paragraphs.FirstOrDefault();
        if (firstPara == null)
        {
            return;
        }

        foreach (var p in paragraphs.Skip(1))
        {
            p.Remove();
        }

        foreach (var portion in firstPara.Portions.ToList())
        {
            portion.Remove();
        }

        int paraIndex = 0;
        foreach (var rawLine in lines)
        {
            if (string.IsNullOrWhiteSpace(rawLine))
            {
                continue;
            }

            var line = rawLine.TrimStart();
            if (!line.StartsWith("- ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var content = line[2..];
            if (paraIndex > 0)
            {
                this.Paragraphs.Add();
            }

            var paragraph = this.Paragraphs[paraIndex];
            foreach (var portion in paragraph.Portions.ToList())
            {
                portion.Remove();
            }

            paragraph.Portions.AddText(content);
            paragraph.Bullet.Type = BulletType.Character;
            paragraph.Bullet.Character = "•";
            paraIndex++;
        }
    }

    private void RenderRegularText(string text)
    {
        var paragraphs = this.Paragraphs.ToList();
        var portionPara = paragraphs.FirstOrDefault(p => p.Portions.Any()) ?? paragraphs.First();

        // Clear other paragraphs
        foreach (var p in paragraphs.Where(p => p != portionPara))
        {
            p.Remove();
        }

        foreach (var portion in portionPara.Portions.ToList())
        {
            portion.Remove();
        }

        const string markdownPattern = @"(\*\*(?<bold>[^\*]+)\*\*)|(?<regular>[^\*]+)";
        var matches = Regex.Matches(
            text, 
            markdownPattern, 
            RegexOptions.Singleline | RegexOptions.IgnoreCase,
            TimeSpan.FromMilliseconds(1000));
        foreach (Match match in matches)
        {
            if (match.Groups["bold"].Success)
            {
                portionPara.Portions.AddText(match.Groups["bold"].Value);
                portionPara.Portions.Last().Font!.IsBold = true;
            }
            else if (match.Groups["regular"].Success)
            {
                portionPara.Portions.AddText(match.Groups["regular"].Value);
                portionPara.Portions.Last().Font!.IsBold = false;
            }
        }

        if (this.AutofitType == AutofitType.Shrink)
        {
            this.ShrinkFont(text);
        }
    }
}