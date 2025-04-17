using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.Texts;

internal sealed class TextBox: ITextBox
{
    private readonly OpenXmlElement textBody;
    private readonly ShapeSize shapeSize;
    private TextVerticalAlignment? vAlignment;

    internal TextBox(OpenXmlElement textBody)
    {
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
            var dontAutofit = aBodyPr.GetFirstChild<A.NoAutoFit>();
            var shrink = aBodyPr.GetFirstChild<A.NormalAutoFit>();
            var resize = aBodyPr.GetFirstChild<A.ShapeAutoFit>();

            switch (value)
            {
                case AutofitType.None:
                    shrink?.Remove();
                    resize?.Remove();
                    dontAutofit = new A.NoAutoFit();
                    aBodyPr.Append(dontAutofit);
                    break;
                case AutofitType.Shrink:
                    dontAutofit?.Remove();
                    resize?.Remove();
                    shrink = new A.NormalAutoFit();
                    aBodyPr.Append(shrink);
                    break;
                case AutofitType.Resize:
                    {
                        dontAutofit?.Remove();
                        shrink?.Remove();
                        resize = new A.ShapeAutoFit();
                        aBodyPr.Append(resize);
                        this.ResizeParentShapeOnDemand();
                        break;
                    }

                default:
                    throw new ArgumentOutOfRangeException(nameof(value), value, null);
            }
        }
    }

    public decimal LeftMargin
    {
        get
        {
                return new LeftRightMargin(this.textBody.GetFirstChild<A.BodyProperties>() !.LeftInset).Value;
        }

        set
        {
            var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
            var emu = new Points(value).AsEmus();
            bodyProperties.LeftInset = new Int32Value((int)emu);
        }
    }

    public decimal RightMargin
    {
        get => new LeftRightMargin(this.textBody.GetFirstChild<A.BodyProperties>() !.RightInset).Value;
        set
        {
            var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
            var emu = new Points(value).AsEmus();
            bodyProperties.RightInset = new Int32Value((int)emu);
        }
    }

    public decimal TopMargin
    {
        get => new TopBottomMargin(this.textBody.GetFirstChild<A.BodyProperties>() !.TopInset).Value;
        set
        {
            var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
            var emu = new Points(value).AsEmus();
            bodyProperties.TopInset = new Int32Value((int)emu);
        }
    }

    public decimal BottomMargin
    {
        get => new TopBottomMargin(this.textBody.GetFirstChild<A.BodyProperties>() !.BottomInset).Value;
        set
        {
            var bodyProperties = this.textBody.GetFirstChild<A.BodyProperties>() !;
            var emu = new Points(value).AsEmus();
            bodyProperties.BottomInset = new Int32Value((int)emu);
        }
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
        if (this.IsList(lines))
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
        var paragraphs = this.Paragraphs.ToList();
        var portionPara = paragraphs.FirstOrDefault(p => p.Portions.Any());
        if (portionPara == null)
        {
            portionPara = paragraphs.First();
            portionPara.Portions.AddText(text);
        }
        else
        {
            var removingParagraphs = paragraphs.Where(p => p != portionPara);
            foreach (var p in removingParagraphs)
            {
                p.Remove();
            }

            portionPara.Text = text;
        }

        if (this.AutofitType == AutofitType.Shrink)
        {
            this.ShrinkText(text, portionPara);
        }

        this.ResizeParentShapeOnDemand();
    }
    
    internal void ResizeParentShapeOnDemand()
    {
        if (this.AutofitType != AutofitType.Resize)
        {
            return;
        }

        var shapeWidthPtCapacity = this.shapeSize.Width - this.LeftMargin - this.RightMargin;
        var shapeHeightPtCapacity = this.shapeSize.Height - this.TopMargin - this.BottomMargin;

        decimal textHeightPx = 0;
        var shapeWidthPxCapacity = new Points(shapeWidthPtCapacity).AsPixels();
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
            var paragraphTextWidthPx = new Text(paragraphText, scFont).WidthPx;
            var paragraphTextHeightPx = new Points(scFont.Size).AsPixels();
            var requiredRowsCount = paragraphTextWidthPx / (decimal)shapeWidthPxCapacity;
            var intRequiredRowsCount = (int)requiredRowsCount;
            var fractionalPart = requiredRowsCount - intRequiredRowsCount;
            if (fractionalPart > 0)
            {
                intRequiredRowsCount++;
            }

            textHeightPx += intRequiredRowsCount * (int)paragraphTextHeightPx;
        }

        this.UpdateShapeHeight(textHeightPx, shapeHeightPtCapacity);
        if (!this.TextWrapped)
        {
            this.UpdateShapeWidth();
        }
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

    private void ShrinkText(string newText, IParagraph paragraph)
    {
        var popularFont = paragraph.Portions.GroupBy(paraPortion => paraPortion.Font!.Size)
            .OrderByDescending(x => x.Count())
            .First().First().Font!;
        var text = new Text(newText, popularFont);
        text.Fit(this.shapeSize.Width, this.shapeSize.Height);
        var internalParagraph = (Paragraph)paragraph;
        internalParagraph.SetFontSize((int)text.FontSize);
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

        var textWidthPx = new Text(longerText, font).WidthPx;
        var leftMarginPx = new Points(this.LeftMargin).AsPixels();
        var rightMarginPx = new Points(this.RightMargin).AsPixels();
        var newWidthPx =
            (int)(textWidthPx *
                  1.4M) // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi. 96/72 = 1.4 
            + leftMarginPx + rightMarginPx;
        var newWidthPt = new Pixels((decimal)newWidthPx).AsPoints();
        this.shapeSize.Width = newWidthPt;
    }

    private void UpdateShapeHeight(decimal textHeightPx, decimal shapeHeightPtCapacity)
    {
        var textHeightPt = new Pixels(textHeightPx).AsPoints();
        var parentShape = this.textBody.Parent!;
        var requiredHeightPt = textHeightPt + this.TopMargin + this.BottomMargin;
        var newHeight = requiredHeightPt + this.TopMargin + this.BottomMargin + this.TopMargin + this.BottomMargin;
        this.shapeSize.Height = newHeight;

        // Raise the shape up by the amount, which is half of the increased offset, like PowerPoint does it
        var position = new Position(parentShape);
        var yOffset = (requiredHeightPt - shapeHeightPtCapacity) / 2;
        position.Y -= yOffset;
    }

    // Detect if the text represents a markdown list
    private bool IsList(string[] lines)
    {
        return lines.Any(l => l.TrimStart().StartsWith("- ", StringComparison.CurrentCulture));
    }

    // Render markdown list items as bullet paragraphs
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

    // Render markdown text with bold formatting
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
        var matches = Regex.Matches(text, markdownPattern, RegexOptions.Singleline | RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(1000));
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
            this.ShrinkText(text, portionPara);
        }
    }
}