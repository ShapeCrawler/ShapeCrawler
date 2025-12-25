using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Texts;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a paragraph.
/// </summary>
public interface IParagraph
{
    /// <summary>
    ///     Gets or sets paragraph text.
    /// </summary>
    string Text { get; set; }

    /// <summary>
    ///     Gets paragraph portion collection.
    /// </summary>
    IParagraphPortions Portions { get; }

    /// <summary>
    ///     Gets bullet.
    /// </summary>
    Bullet Bullet { get; }

    /// <summary>
    ///     Gets or sets the text horizontal alignment.
    /// </summary>
    TextHorizontalAlignment HorizontalAlignment { get; set; }

    /// <summary>
    ///     Gets spacing.
    /// </summary>
    ISpacing Spacing { get; }

    /// <summary>
    ///     Gets font color.
    /// </summary>
    string FontColor { get; }

    /// <summary>
    ///    Gets paragraph left margin in points.
    /// </summary>
    decimal LeftMargin { get; }

    /// <summary>
    ///     Gets or sets paragraph indent level.
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    ///     Gets or sets paragraph first line indent in points.
    /// </summary>
    decimal FirstLineIndent { get; set; }

    /// <summary>
    ///     Finds and replaces text.
    /// </summary>
    void ReplaceText(string oldValue, string newValue);

    /// <summary>
    ///     Removes paragraph.
    /// </summary>
    void Remove();

    /// <summary>
    ///     Sets font size in points.
    /// </summary>
    void SetFontSize(int fontSize);

    /// <summary>
    ///     Sets font name.
    /// </summary>
    void SetFontName(string fontName);

    /// <summary>
    ///     Sets font color.
    /// </summary>
    void SetFontColor(string colorHex);

    /// <summary>
    ///    Sets paragraph left margin in points.
    /// </summary>
    void SetLeftMargin(decimal points);
}

internal sealed class Paragraph : IParagraph
{
    private readonly Lazy<Bullet> bullet;
    private readonly SCAParagraph scAParagraph;
    private readonly A.Paragraph aParagraph;
    private readonly ParagraphPortions portions;
    private TextHorizontalAlignment? alignment;

    internal Paragraph(A.Paragraph aParagraph)
        : this(aParagraph, new SCAParagraph(aParagraph))
    {
    }

    private Paragraph(A.Paragraph aParagraph, SCAParagraph scAParagraph)
    {
        this.aParagraph = aParagraph;
        this.scAParagraph = scAParagraph;
        this.aParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        this.bullet = new Lazy<Bullet>(this.GetBullet);
        this.portions = new ParagraphPortions(this.aParagraph);
    }

    public string Text
    {
        get
        {
            if (this.portions.Count == 0)
            {
                return string.Empty;
            }

            return this.portions.Select(portion => portion.Text).Aggregate((result, next) => result + next)!;
        }

        set
        {
            if (!this.portions.Any())
            {
                this.portions.AddText(" ");
            }

            var removingRuns = this.aParagraph.OfType<A.Run>().Skip(1); // to preserve text formatting
            var removingBreaks = this.aParagraph.OfType<A.Break>();
            foreach (var removing in removingRuns.ToArray())
            {
                removing.Remove();
            }

            foreach (var removing in removingBreaks.ToList())
            {
                removing.Remove();
            }

#if NETSTANDARD2_0
            var textLines = value.Split([Environment.NewLine], StringSplitOptions.None);
#else
            var textLines = value.Split(Environment.NewLine);
#endif
            var mainRun = this.aParagraph.GetFirstChild<A.Run>()!;
            mainRun.Text!.Text = textLines[0];

            foreach (var textLine in textLines.Skip(1))
            {
                if (!string.IsNullOrEmpty(textLine))
                {
                    this.portions.AddLineBreak();
                    this.portions.AddText(textLine);
                }
                else
                {
                    this.portions.AddLineBreak();
                }
            }

            var textBody = this.aParagraph.Parent!;
            var textBox = new DrawingTextBox(new TextBoxMargins(textBody), textBody);
            textBox.ResizeParentShapeOnDemand();
        }
    }

    public IParagraphPortions Portions => this.portions;

    public Bullet Bullet => this.bullet.Value;

    public TextHorizontalAlignment HorizontalAlignment
    {
        get
        {
            if (this.alignment.HasValue)
            {
                return this.alignment.Value;
            }

            var calculatedAlignment = new ParagraphHorizontalAlignment(this.aParagraph).ValueOrNull();
            this.alignment = calculatedAlignment ?? TextHorizontalAlignment.Left;
            return this.alignment.Value;
        }
        set => this.SetAlignment(value);
    }

    public int IndentLevel
    {
        get => this.scAParagraph.GetIndentLevel();
        set => this.scAParagraph.UpdateIndentLevel(value);
    }

    public ISpacing Spacing => this.GetSpacing();

    public string FontColor
    {
        get
        {
            if (this.Portions.Count == 0)
            {
                return string.Empty;
            }

            return this.Portions[0].Font!.Color.Hex;
        }
    }

    public decimal LeftMargin
    {
        get
        {
            var leftMargin = this.aParagraph.ParagraphProperties!.LeftMargin;
            if (leftMargin is not null)
            {
                return new Emus(leftMargin.Value).AsPoints();
            }

            return this.IndentationFromStylesOrDefault("marL");
        }

        set
        {
            var leftMarginEmu = (int)new Points(value).AsEmus();
            this.aParagraph.ParagraphProperties!.LeftMargin = new Int32Value(leftMarginEmu);
        }
    }

    public decimal FirstLineIndent
    {
        get
        {
            var indent = this.aParagraph.ParagraphProperties!.Indent;
            if (indent is not null)
            {
                return new Emus(indent.Value).AsPoints();
            }

            return this.IndentationFromStylesOrDefault("indent");
        }

        set
        {
            var indentEmu = (int)new Points(value).AsEmus();
            this.aParagraph.ParagraphProperties!.Indent = new Int32Value(indentEmu);
        }
    }

    public void ReplaceText(string oldValue, string newValue)
    {
        foreach (var portion in this.portions.Where(portion => portion is not ParagraphLineBreak))
        {
            portion.Text = portion.Text.Replace(oldValue, newValue);
        }

        if (this.Text.Contains(oldValue))
        {
            this.Text = this.Text.Replace(oldValue, newValue);
        }
    }

    public void Remove() => this.aParagraph.Remove();

    public void SetFontSize(int fontSize)
    {
        foreach (var portion in this.portions)
        {
            portion.Font!.Size = fontSize;
        }
    }

    public void SetFontName(string fontName)
    {
        foreach (var portion in this.Portions)
        {
            portion.Font!.LatinName = fontName;
        }
    }

    public void SetFontColor(string colorHex)
    {
        if (!this.Portions.Any())
        {
            this.Portions.AddText(" ");
        }

        foreach (var portion in this.Portions)
        {
            portion.Font!.Color.Set(colorHex);
        }

        // Also set on EndParagraphRunProperties so newly typed text inherits the color
        var endParaRPr = this.aParagraph.GetFirstChild<A.EndParagraphRunProperties>();
        if (endParaRPr != null)
        {
            colorHex = colorHex.StartsWith("#", StringComparison.Ordinal) ? colorHex[1..] : colorHex;
            if (colorHex.Length == 8)
            {
                colorHex = colorHex[..6];
            }

            endParaRPr.GetFirstChild<A.SolidFill>()?.Remove();
            var solidFill = new A.SolidFill(new A.RgbColorModelHex { Val = colorHex });
            endParaRPr.InsertAt(solidFill, 0);
        }
    }

    public void SetLeftMargin(decimal points)
    {
        this.LeftMargin = points;
    }

    private static long? EmusAttributeFromIndentStylesOrNull(
        OpenXmlCompositeElement? openXmlCompositeElement,
        int indentLevel,
        string attributeLocalName)
    {
        if (openXmlCompositeElement is null)
        {
            return null;
        }

        foreach (var levelProperties in openXmlCompositeElement.Elements()
                     .Where(e => e.LocalName.StartsWith("lvl", StringComparison.Ordinal)))
        {
            var level = ExtractLevelNumberOrZero(levelProperties.LocalName);
            if (level != indentLevel)
            {
                continue;
            }

            var attributeValue = levelProperties
                .GetAttributes()
                .Where(a => a.LocalName == attributeLocalName)
                .Select(a => a.Value)
                .FirstOrDefault();

            if (string.IsNullOrWhiteSpace(attributeValue))
            {
                return null;
            }

            return long.TryParse(attributeValue, out var emus) ? emus : null;
        }

        return null;
    }

    private static int ExtractLevelNumberOrZero(string localName)
    {
        if (localName.Length < 4)
        {
            return 0;
        }

        var levelChar = localName[3];
        return levelChar >= '0' && levelChar <= '9' ? levelChar - '0' : 0;
    }

    private decimal IndentationFromStylesOrDefault(string attributeLocalName)
    {
        var indentLevel = this.IndentLevel;

        var listStyle = this.aParagraph.Parent?.GetFirstChild<A.ListStyle>();
        var listStyleEmus = EmusAttributeFromIndentStylesOrNull(listStyle, indentLevel, attributeLocalName);
        if (listStyleEmus.HasValue)
        {
            return new Emus(listStyleEmus.Value).AsPoints();
        }

        var defaultTextStyle = this.DefaultTextStyleOrNull();
        var defaultTextStyleEmus = EmusAttributeFromIndentStylesOrNull(defaultTextStyle, indentLevel, attributeLocalName);
        if (defaultTextStyleEmus.HasValue)
        {
            return new Emus(defaultTextStyleEmus.Value).AsPoints();
        }

        return 0;
    }

    private OpenXmlCompositeElement? DefaultTextStyleOrNull()
    {
        var openXmlPartOrNull = this.aParagraph.Ancestors<OpenXmlPartRootElement>().FirstOrDefault()?.OpenXmlPart;
        if (openXmlPartOrNull?.OpenXmlPackage is not PresentationDocument presDocument)
        {
            return null;
        }

        return presDocument.PresentationPart?.Presentation.DefaultTextStyle;
    }

    private ISpacing GetSpacing() => new Spacing(this.aParagraph);

    private Bullet GetBullet() => new(this.aParagraph.ParagraphProperties!);

    private void SetAlignment(TextHorizontalAlignment alignmentValue)
    {
        var aTextAlignmentTypeValue = alignmentValue switch
        {
            TextHorizontalAlignment.Left => A.TextAlignmentTypeValues.Left,
            TextHorizontalAlignment.Center => A.TextAlignmentTypeValues.Center,
            TextHorizontalAlignment.Right => A.TextAlignmentTypeValues.Right,
            TextHorizontalAlignment.Justify => A.TextAlignmentTypeValues.Justified,
            _ => throw new ArgumentOutOfRangeException(nameof(alignmentValue))
        };

        if (this.aParagraph.ParagraphProperties == null)
        {
            this.aParagraph.ParagraphProperties = new A.ParagraphProperties
            {
                Alignment = new EnumValue<A.TextAlignmentTypeValues>(aTextAlignmentTypeValue)
            };
        }
        else
        {
            this.aParagraph.ParagraphProperties.Alignment =
                new EnumValue<A.TextAlignmentTypeValues>(aTextAlignmentTypeValue);
        }

        this.alignment = alignmentValue;
    }
}