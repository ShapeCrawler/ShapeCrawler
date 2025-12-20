using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Texts;
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
    ///     Gets or sets paragraph indent level.
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    ///     Gets spacing.
    /// </summary>
    ISpacing Spacing { get; }
    
    /// <summary>
    ///     Gets font color.
    /// </summary>
    string FontColor { get; }

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

            return this.portions.Select(portion => portion.Text).Aggregate((result, next) => result + next) !;
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
            var mainRun = this.aParagraph.GetFirstChild<A.Run>() !;
            mainRun.Text!.Text = textLines.First();

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
        var leftMarginEmu = (int)new ShapeCrawler.Units.Points(points).AsEmus();
        this.aParagraph.ParagraphProperties!.LeftMargin = new Int32Value(leftMarginEmu);
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