using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Texts;
using ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
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
    ///     Gets collection of paragraph portions.
    /// </summary>
    IParagraphPortions Portions { get; }

    /// <summary>
    ///     Gets paragraph bullet if bullet exist, otherwise <see langword="null"/>.
    /// </summary>
    Bullet Bullet { get; }

    /// <summary>
    ///     Gets or sets the text alignment.
    /// </summary>
    TextAlignment Alignment { get; set; }

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

    void Remove();
}

internal sealed class Paragraph : IParagraph
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly Lazy<Bullet> bullet;
    private TextAlignment? alignment;
    private readonly AParagraphWrap aParagraphWrap;

    internal Paragraph(TypedOpenXmlPart sdkTypedOpenXmlPart, A.Paragraph aParagraph)
        : this(sdkTypedOpenXmlPart, aParagraph, new AParagraphWrap(aParagraph))
    {
    }

    private Paragraph(TypedOpenXmlPart sdkTypedOpenXmlPart, A.Paragraph aParagraph, AParagraphWrap aParagraphWrap)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.AParagraph = aParagraph;
        this.aParagraphWrap = aParagraphWrap;
        this.AParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        this.bullet = new Lazy<Bullet>(this.GetBullet);
        this.Portions = new ParagraphPortions(sdkTypedOpenXmlPart, this.AParagraph);
    }

    public bool IsRemoved { get; set; }

    public string Text
    {
        get => this.ParseText();
        set
        {
            if (!this.Portions.Any())
            {
                this.Portions.AddText(" ");
            }

            // To set a paragraph text we use a single portion which is the first paragraph portion.
            var baseARun = this.AParagraph.GetFirstChild<A.Run>()!;
            foreach (var removingRun in this.AParagraph.OfType<A.Run>().Where(run => run != baseARun))
            {
                removingRun.Remove();
            }

#if NETSTANDARD2_0
        var textLines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
#else
            var textLines = value.Split(Environment.NewLine);
#endif

            var basePortion = new TextParagraphPortion(this.sdkTypedOpenXmlPart, baseARun);
            basePortion.Text = textLines.First();

            foreach (var textLine in textLines.Skip(1))
            {
                if (!string.IsNullOrEmpty(textLine))
                {
                    ((ParagraphPortions)this.Portions).AddNewLine();
                    this.Portions.AddText(textLine);
                }
                else
                {
                    ((ParagraphPortions)this.Portions).AddNewLine();
                }
            }

            // Resize
            var sdkTextBody = this.AParagraph.Parent!;
            var textFrame = new TextFrame(this.sdkTypedOpenXmlPart, sdkTextBody);
            textFrame.ResizeParentShape();
        }
    }

    public IParagraphPortions Portions { get; }

    public Bullet Bullet => this.bullet.Value;

    public TextAlignment Alignment
    {
        get => this.ParseAlignment();
        set => this.SetAlignment(value);
    }

    public int IndentLevel
    {
        get => this.aParagraphWrap.IndentLevel();
        set => this.aParagraphWrap.UpdateIndentLevel(value);
    }

    public ISpacing Spacing => this.GetSpacing();
    internal A.Paragraph AParagraph { get; }

    public void SetFontSize(int fontSize)
    {
        foreach (var portion in this.Portions)
        {
            portion.Font.Size = fontSize;
        }
    }

    public void ReplaceText(string oldValue, string newValue)
    {
        foreach (var portion in this.Portions)
        {
            portion.Text = portion.Text!.Replace(oldValue, newValue);
        }

        if (this.Text.Contains(oldValue))
        {
            this.Text = this.Text.Replace(oldValue, newValue);
        }
    }

    public void Remove() => this.AParagraph.Remove();
    private ISpacing GetSpacing() => new Spacing(this, this.AParagraph);

    private Bullet GetBullet() => new Bullet(this.AParagraph.ParagraphProperties!);

    private string ParseText()
    {
        if (this.Portions.Count == 0)
        {
            return string.Empty;
        }

        return this.Portions.Select(portion => portion.Text).Aggregate((result, next) => result + next) !;
    }

    private void SetAlignment(TextAlignment alignmentValue)
    {
        var aTextAlignmentTypeValue = alignmentValue switch
        {
            TextAlignment.Left => A.TextAlignmentTypeValues.Left,
            TextAlignment.Center => A.TextAlignmentTypeValues.Center,
            TextAlignment.Right => A.TextAlignmentTypeValues.Right,
            TextAlignment.Justify => A.TextAlignmentTypeValues.Justified,
            _ => throw new ArgumentOutOfRangeException(nameof(alignmentValue))
        };

        if (this.AParagraph.ParagraphProperties == null)
        {
            this.AParagraph.ParagraphProperties = new A.ParagraphProperties
            {
                Alignment = new EnumValue<A.TextAlignmentTypeValues>(aTextAlignmentTypeValue)
            };
        }
        else
        {
            this.AParagraph.ParagraphProperties.Alignment =
                new EnumValue<A.TextAlignmentTypeValues>(aTextAlignmentTypeValue);
        }

        this.alignment = alignmentValue;
    }

    private TextAlignment ParseAlignment()
    {
        if (this.alignment.HasValue)
        {
            return this.alignment.Value;
        }

        var aTextAlignmentType = this.AParagraph.ParagraphProperties?.Alignment!;
        if (aTextAlignmentType == null)
        {
            return TextAlignment.Left;
        }

        this.alignment = aTextAlignmentType.Value switch
        {
            A.TextAlignmentTypeValues.Center => TextAlignment.Center,
            A.TextAlignmentTypeValues.Right => TextAlignment.Right,
            A.TextAlignmentTypeValues.Justified => TextAlignment.Justify,
            _ => TextAlignment.Left
        };

        return this.alignment.Value;
    }
}