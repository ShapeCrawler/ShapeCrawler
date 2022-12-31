using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Collections;
using ShapeCrawler.Factories;
using ShapeCrawler.Shared;
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
    IPortionCollection Portions { get; }

    /// <summary>
    ///     Gets paragraph bullet if bullet exist, otherwise <see langword="null"/>
    /// </summary>
    SCBullet Bullet { get; }

    /// <summary>
    ///     Gets or sets the text alignment.
    /// </summary>
    SCTextAlignment Alignment { get; set; }

    /// <summary>
    ///     Gets paragraph's indent level.
    /// </summary>
    int IndentLevel { get; }

    /// <summary>
    ///     Gets spacing.
    /// </summary>
    ISpacing Spacing { get; }

    /// <summary>
    ///     Gets font.
    /// </summary>
    IFont Font { get; }

    /// <summary>
    ///     Adds new text portion in paragraph.
    /// </summary>
    void AddPortion(string text);

    /// <summary>
    ///     Finds and replaces text.
    /// </summary>
    void ReplaceText(string oldValue, string newValue);
}

internal sealed class SCParagraph : IParagraph
{
    private readonly Lazy<SCBullet> bullet;
    private readonly ResettableLazy<PortionCollection> portions;
    private SCTextAlignment? alignment;

    internal SCParagraph(A.Paragraph aParagraph, TextFrame textBox)
    {
        this.AParagraph = aParagraph;
        this.AParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        this.Level = this.GetIndentLevel();
        this.bullet = new Lazy<SCBullet>(this.GetBullet);
        this.ParentTextFrame = textBox;
        this.portions = new ResettableLazy<PortionCollection>(() => new PortionCollection(this.AParagraph, this));
    }

    internal event Action? TextChanged;

    public bool IsRemoved { get; set; }

    public string Text
    {
        get => this.GetText();
        set => this.SetText(value);
    }

    public IPortionCollection Portions => this.portions.Value;

    public SCBullet Bullet => this.bullet.Value;

    public SCTextAlignment Alignment
    {
        get => this.GetAlignment();
        set => this.SetAlignment(value);
    }

    public int IndentLevel => this.GetIndentLevel();

    public ISpacing Spacing => this.GetSpacing();

    public IFont Font => this.GetFont();

    internal TextFrame ParentTextFrame { get; }

    internal A.Paragraph AParagraph { get; }

    internal int Level { get; }

    public void SetFontSize(int fontSize)
    {
        foreach (var portion in this.Portions)
        {
            portion.Font.Size = fontSize;
        }
    }

    public void AddPortion(string text)
    {
        if (text == string.Empty)
        {
            return;
        }

        var lastPortion = this.portions.Value.LastOrDefault() as SCPortion;
        OpenXmlElement aRun;
        OpenXmlElement? lastARunOrABreak = null;
        if (lastPortion == null)
        {
            var aRunBuilder = new ARunBuilder();
            aRun = aRunBuilder.Build();
        }
        else
        {
            aRun = lastPortion.AText.Parent!;
            lastARunOrABreak = this.AParagraph.Last(p => p is A.Run or A.Break);
            if (lastARunOrABreak is not A.Break && this.Text.EndsWith(Environment.NewLine, StringComparison.Ordinal))
            {
                AddBreak(ref lastARunOrABreak);
            }
        }

        var textLines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
        if (lastPortion?.Text == string.Empty)
        {
            lastPortion.Text = textLines[0];
        }
        else
        {
            AddText(ref lastARunOrABreak, aRun, textLines[0], this.AParagraph);
        }

        for (var i = 1; i < textLines.Length; i++)
        {
            AddBreak(ref lastARunOrABreak!);
            if (textLines[i] != string.Empty)
            {
                AddText(ref lastARunOrABreak, aRun, textLines[i], this.AParagraph);
            }
        }

        this.portions.Reset();
    }

    public void ReplaceText(string oldValue, string newValue)
    {
        foreach (var portion in this.Portions)
        {
            portion.Text = portion.Text.Replace(oldValue, newValue);
        }

        if (this.Text.Contains(oldValue))
        {
            this.Text = this.Text.Replace(oldValue, newValue);
        }
    }

    private static void AddBreak(ref OpenXmlElement lastElement)
    {
        lastElement = lastElement.InsertAfterSelf(new A.Break());
    }
    
    private static void AddText(
        ref OpenXmlElement? lastElement,
        OpenXmlElement aTextParent,
        string text,
        A.Paragraph aParagraph)
    {
        var newARun = (A.Run)aTextParent.CloneNode(true);
        newARun.Text!.Text = text;
        if (lastElement == null)
        {
            aParagraph.InsertAt(newARun, 0);
        }
        else
        {
            lastElement = lastElement.InsertAfterSelf(newARun);
        }
    }
    
    private IFont GetFont()
    {
        throw new NotImplementedException();
    }
    
    private ISpacing GetSpacing()
    {
        return new SCSpacing(this, this.AParagraph);
    }

    private SCBullet GetBullet()
    {
        return new SCBullet(this.AParagraph.ParagraphProperties!);
    }

    private string GetText()
    {
        if (this.Portions.Count == 0)
        {
            return string.Empty;
        }

        return this.Portions.Select(portion => portion.Text).Aggregate((result, next) => result + next);
    }

    private int GetIndentLevel()
    {
        var level = this.AParagraph.ParagraphProperties!.Level;
        if (level is null)
        {
            return 1; // it is default indent level
        }

        return level + 1;
    }

    private void SetText(string text)
    {
        if (this.portions.Value.Count == 0)
        {
            this.AddPortion(" ");
        }

        // to set a paragraph text we use a single portion which is the first paragraph portion.
        var removingPortions = this.Portions.Skip(1).ToList();
        this.Portions.Remove(removingPortions);

        var basePortion = (SCPortion)this.portions.Value.Single();
        if (text.Contains(Environment.NewLine))
        {
            basePortion.Text = string.Empty;
            this.AddPortion(text);
        }
        else
        {
            basePortion.Text = text;
        }

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

    private SCTextAlignment GetAlignment()
    {
        if (this.alignment.HasValue)
        {
            return this.alignment.Value;
        }

        var shape = this.ParentTextFrame.TextFrameContainer.Shape;
        var placeholder = shape.Placeholder;

        var aTextAlignmentType = this.AParagraph.ParagraphProperties?.Alignment!;
        if (aTextAlignmentType == null)
        {
            if (placeholder is { Type: SCPlaceholderType.Title })
            {
                this.alignment = SCTextAlignment.Left;
                return this.alignment.Value;
            }

            if (placeholder is { Type: SCPlaceholderType.CenteredTitle })
            {
                this.alignment = SCTextAlignment.Center;
                return this.alignment.Value;
            }

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