using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Paragraphs;
using ShapeCrawler.Shapes;
using ShapeCrawler.Texts;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

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
    ///     Gets the collection of paragraph portions.
    /// </summary>
    IParagraphPortions Portions { get; }

    /// <summary>
    ///     Gets paragraph's bullet. Returns <see langword="null"/> if bullet doesn't exist.
    /// </summary>
    Bullet Bullet { get; }

    /// <summary>
    ///     Gets or sets the text horizontal alignment.
    /// </summary>
    TextHorizontalAlignment HorizontalAlignment { get; set; }

    /// <summary>
    ///     Gets or sets paragraph's indent level.
    /// </summary>
    int IndentLevel { get; set; }

    /// <summary>
    ///     Gets spacing.
    /// </summary>
    ISpacing Spacing { get; }

    /// <summary>
    ///     Finds and replaces text.
    /// </summary>
    void ReplaceText(string oldValue, string newValue);

    /// <summary>
    ///     Removes paragraph.
    /// </summary>
    void Remove();
}

internal sealed class Paragraph : IParagraph
{
    private readonly Lazy<Bullet> bullet;
    private readonly SCAParagraph scAParagraph;
    private readonly A.Paragraph aParagraph;
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
        this.Portions = new ParagraphPortions(this.aParagraph);
    }

    public string Text
    {
        get
        {
            if (this.Portions.Count == 0)
            {
                return string.Empty;
            }

            return this.Portions.Select(portion => portion.Text).Aggregate((result, next) => result + next) !;
        }

        set
        {
            if (!this.Portions.Any())
            {
                this.Portions.AddText(" ");
            }

            var removingRuns = this.aParagraph.OfType<A.Run>().Skip(1); // to preserve text formatting
            var removingBreaks = this.aParagraph.OfType<A.Break>();
            foreach (var removing in removingRuns.ToList())
            {
                removing.Remove();
            }

            foreach (var removing in removingBreaks.ToList())
            {
                removing.Remove();
            }

#if NETSTANDARD2_0
            var textLines = value.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
#else
            var textLines = value.Split(Environment.NewLine);
#endif
            var mainRun = this.aParagraph.GetFirstChild<A.Run>() !;
            mainRun.Text!.Text = textLines.First();

            foreach (var textLine in textLines.Skip(1))
            {
                if (!string.IsNullOrEmpty(textLine))
                {
                    ((ParagraphPortions)this.Portions).AddLineBreak();
                    this.Portions.AddText(textLine);
                }
                else
                {
                    ((ParagraphPortions)this.Portions).AddLineBreak();
                }
            }

            // Resize
            var textBody = this.aParagraph.Parent!;
            var textBox = new TextBox(textBody);
            textBox.ResizeParentShapeOnDemand();
        }
    }

    public IParagraphPortions Portions { get; }

    public Bullet Bullet => this.bullet.Value;

    public TextHorizontalAlignment HorizontalAlignment
    {
        get
        {
            if (this.alignment.HasValue)
            {
                return this.alignment.Value;
            }

            var aTextAlignmentType = this.aParagraph.ParagraphProperties?.Alignment;
            if (aTextAlignmentType == null)
            {
                var parentShape = new AutoShape(this.aParagraph.Ancestors<P.Shape>().First());
                if (parentShape.PlaceholderType == PlaceholderType.CenteredTitle)
                {
                    return TextHorizontalAlignment.Center;
                }
            }

            if (aTextAlignmentType is null)
            {
                return TextHorizontalAlignment.Center;
            }

            if (aTextAlignmentType!.Value == A.TextAlignmentTypeValues.Center)
            {
                this.alignment = TextHorizontalAlignment.Center;
            }
            else if (aTextAlignmentType!.Value == A.TextAlignmentTypeValues.Right)
            {
                this.alignment = TextHorizontalAlignment.Right;
            }
            else if (aTextAlignmentType!.Value == A.TextAlignmentTypeValues.Justified)
            {
                this.alignment = TextHorizontalAlignment.Justify;
            }
            else
            {
                this.alignment = TextHorizontalAlignment.Left;
            }

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

    public void ReplaceText(string oldValue, string newValue)
    {
        foreach (var portion in this.Portions.Where(portion => portion is not ParagraphLineBreak))
        {
            portion.Text = portion.Text.Replace(oldValue, newValue);
        }

        if (this.Text.Contains(oldValue))
        {
            this.Text = this.Text.Replace(oldValue, newValue);
        }
    }

    public void Remove() => this.aParagraph.Remove();

    internal void SetFontSize(int fontSize)
    {
        foreach (var portion in this.Portions)
        {
            portion.Font!.Size = fontSize;
        }
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