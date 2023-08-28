using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;

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
    IParagraphPortionCollection Portions { get; }

    /// <summary>
    ///     Gets paragraph bullet if bullet exist, otherwise <see langword="null"/>.
    /// </summary>
    SCBullet Bullet { get; }

    /// <summary>
    ///     Gets or sets the text alignment.
    /// </summary>
    SCTextAlignment Alignment { get; set; }

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
}

internal sealed class SlideParagraph : IParagraph
{
    private readonly Lazy<SCBullet> bullet;
    private readonly ResetableLazy<SlideParagraphPortions> portions;
    private SCTextAlignment? alignment;
    private readonly SlidePart sdkSlidePart;
    private readonly SDKAParagraphWrap sdkaParagraphWrap;

    internal SlideParagraph(SlidePart sdkSlidePart, A.Paragraph aParagraph)
        : this(sdkSlidePart, aParagraph, new SDKAParagraphWrap(aParagraph))
    {
    }

    private SlideParagraph(SlidePart sdkSlidePart, A.Paragraph aParagraph, SDKAParagraphWrap sdkaParagraphWrap)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.AParagraph = aParagraph;
        this.sdkaParagraphWrap = sdkaParagraphWrap;
        this.AParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        this.bullet = new Lazy<SCBullet>(this.GetBullet);
        this.portions = new ResetableLazy<SlideParagraphPortions>(() => new SlideParagraphPortions(this.sdkSlidePart,this.AParagraph));
    }

    internal event Action? TextChanged;

    public bool IsRemoved { get; set; }

    public string Text
    {
        get => this.ParseText();
        set => this.SetText(value);
    }

    public IParagraphPortionCollection Portions => this.portions.Value;

    public SCBullet Bullet => this.bullet.Value;

    public SCTextAlignment Alignment
    {
        get => this.ParseAlignment();
        set => this.SetAlignment(value);
    }

    public int IndentLevel
    {
        get => this.sdkaParagraphWrap.IndentLevel();
        set => this.sdkaParagraphWrap.UpdateIndentLevel(value);
    }

    public ISpacing Spacing => this.GetSpacing();

    internal A.Paragraph AParagraph { get; }

    public void SetFontSize(int fontSize)
    {
        foreach (var portion in this.Portions)
        {
            portion.Font!.Size = fontSize;
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

    private ISpacing GetSpacing()
    {
        return new SCSpacing(this, this.AParagraph);
    }

    private SCBullet GetBullet()
    {
        return new SCBullet(this.AParagraph.ParagraphProperties!);
    }

    private string ParseText()
    {
        if (this.Portions.Count == 0)
        {
            return string.Empty;
        }

        return this.Portions.Select(portion => portion.Text).Aggregate((result, next) => result + next) !;
    }

    private void SetText(string text)
    {
        if (!this.portions.Value.Any())
        {
            this.portions.Value.AddText(" ");
        }

        // To set a paragraph text we use a single portion which is the first paragraph portion.
        var basePortion = this.portions.Value.OfType<TextParagraphPortion>().First();
        var removingPortions = this.portions.Value.Where(p => p != basePortion).ToList();
        this.portions.Value.Remove(removingPortions);

#if NETSTANDARD2_0
        var textLines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
#else
        var textLines = text.Split(Environment.NewLine);
#endif

        basePortion.Text = textLines.First();

        foreach (var textLine in textLines.Skip(1))
        {
            if (!string.IsNullOrEmpty(textLine))
            {
                this.portions.Value.AddNewLine();
                this.portions.Value.AddText(textLine);
            }
            else
            {
                this.portions.Value.AddNewLine();
            }
        }

        this.portions.Reset();
        this.TextChanged?.Invoke();
    }

    private void SetAlignment(SCTextAlignment alignmentValue)
    {
        var aTextAlignmentTypeValue = alignmentValue switch
        {
            SCTextAlignment.Left => A.TextAlignmentTypeValues.Left,
            SCTextAlignment.Center => A.TextAlignmentTypeValues.Center,
            SCTextAlignment.Right => A.TextAlignmentTypeValues.Right,
            SCTextAlignment.Justify => A.TextAlignmentTypeValues.Justified,
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

    private SCTextAlignment ParseAlignment()
    {
        if (this.alignment.HasValue)
        {
            return this.alignment.Value;
        }

        var aTextAlignmentType = this.AParagraph.ParagraphProperties?.Alignment!;
        if (aTextAlignmentType == null)
        {
            return SCTextAlignment.Left;
        }

        this.alignment = aTextAlignmentType.Value switch
        {
            A.TextAlignmentTypeValues.Center => SCTextAlignment.Center,
            A.TextAlignmentTypeValues.Right => SCTextAlignment.Right,
            A.TextAlignmentTypeValues.Justified => SCTextAlignment.Justify,
            _ => SCTextAlignment.Left
        };

        return this.alignment.Value;
    }
}