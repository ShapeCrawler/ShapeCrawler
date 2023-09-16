using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using ShapeCrawler.Wrappers;
using SkiaSharp;
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
    private SCTextAlignment? alignment;
    private readonly SlidePart sdkSlidePart;
    private readonly SdkAParagraph sdkAParagraph;

    internal SlideParagraph(SlidePart sdkSlidePart, A.Paragraph aParagraph)
        : this(sdkSlidePart, aParagraph, new SdkAParagraph(aParagraph))
    {
    }

    private SlideParagraph(SlidePart sdkSlidePart, A.Paragraph aParagraph, SdkAParagraph sdkAParagraph)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.AParagraph = aParagraph;
        this.sdkAParagraph = sdkAParagraph;
        this.AParagraph.ParagraphProperties ??= new A.ParagraphProperties();
        this.bullet = new Lazy<SCBullet>(this.GetBullet);
        this.Portions = new SlideParagraphPortions(this.sdkSlidePart,this.AParagraph); 
    }

    public bool IsRemoved { get; set; }

    public string Text
    {
        get => this.ParseText();
        set => this.UpdateText(value);
    }

    public IParagraphPortions Portions { get; }

    public SCBullet Bullet => this.bullet.Value;

    public SCTextAlignment Alignment
    {
        get => this.ParseAlignment();
        set => this.SetAlignment(value);
    }

    public int IndentLevel
    {
        get => this.sdkAParagraph.IndentLevel();
        set => this.sdkAParagraph.UpdateIndentLevel(value);
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

    private ISpacing GetSpacing() => new Spacing(this, this.AParagraph);

    private SCBullet GetBullet()=> new SCBullet(this.AParagraph.ParagraphProperties!);

    private string ParseText()
    {
        if (this.Portions.Count == 0)
        {
            return string.Empty;
        }

        return this.Portions.Select(portion => portion.Text).Aggregate((result, next) => result + next) !;
    }

    private void UpdateText(string text)
    {
        if (!this.Portions.Any())
        {
            this.Portions.AddText(" ");
        }

        // To set a paragraph text we use a single portion which is the first paragraph portion.
        // var basePortion = this.Portions.OfType<TextParagraphPortion>().First();
        // var removingPortions = this.Portions.Where(p => p != basePortion).ToList();
        // this.Portions.Remove(removingPortions);
        var baseARun = this.AParagraph.GetFirstChild<A.Run>()!;
        foreach (var removingRun in this.AParagraph.OfType<A.Run>().Where(run => run != baseARun))
        {
            removingRun.Remove();
        }

#if NETSTANDARD2_0
        var textLines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
#else
        var textLines = text.Split(Environment.NewLine);
#endif

        var basePortion = new SlideTextParagraphPortion(this.sdkSlidePart, baseARun);
        basePortion.Text = textLines.First();

        foreach (var textLine in textLines.Skip(1))
        {
            if (!string.IsNullOrEmpty(textLine))
            {
                ((SlideParagraphPortions)this.Portions).AddNewLine();
                this.Portions.AddText(textLine);
            }
            else
            {
                ((SlideParagraphPortions)this.Portions).AddNewLine();
            }
        }
        
        // Resize
        var pTextBody = (P.TextBody)this.AParagraph.Parent!;
        var textFrame = new TextFrame(this.sdkSlidePart, pTextBody);
        if (textFrame.AutofitType != SCAutofitType.Resize)
        {
            return;
        }

        var baseParagraph = textFrame.Paragraphs.First();
        var popularPortion = baseParagraph.Portions.OfType<SlideTextParagraphPortion>().GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var paint = new SKPaint();
        var fontSize = font!.Size;
        paint.TextSize = fontSize;
        paint.Typeface = SKTypeface.FromFamilyName(font.LatinName);
        paint.IsAntialias = true;

        var lMarginPixel = UnitConverter.CentimeterToPixel(textFrame.LeftMargin);
        var rMarginPixel = UnitConverter.CentimeterToPixel(textFrame.RightMargin);
        var tMarginPixel = UnitConverter.CentimeterToPixel(textFrame.TopMargin);
        var bMarginPixel = UnitConverter.CentimeterToPixel(textFrame.BottomMargin);

        var textRect = default(SKRect);
        // var text = textFrame.Text;
        paint.MeasureText(text, ref textRect);
        var textWidth = textRect.Width;
        var textHeight = paint.TextSize;
        var shapeSize = new ShapeSize(pTextBody.Parent!);
        var currentBlockWidth = shapeSize.Width() - lMarginPixel - rMarginPixel;
        var currentBlockHeight = shapeSize.Height() - tMarginPixel - bMarginPixel;

        this.UpdateHeight(textWidth, currentBlockWidth, textHeight, tMarginPixel, bMarginPixel, currentBlockHeight, pTextBody.Parent!);
        this.UpdateWidthIfNeed(paint, lMarginPixel, rMarginPixel, textFrame, pTextBody.Parent!);
    }
    
    private void UpdateWidthIfNeed(SKPaint paint, int lMarginPixel, int rMarginPixel, TextFrame textFrame, OpenXmlElement parent)
    {
        if (!textFrame.TextWrapped)
        {
            var longerText = textFrame.Paragraphs
                .Select(x => new { x.Text, x.Text.Length })
                .OrderByDescending(x => x.Length)
                .First().Text;
            var paraTextRect = default(SKRect);
            var widthInPixels = paint.MeasureText(longerText, ref paraTextRect);
            // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi.
            // 96/72=1.4
            const double Scale = 1.4;
            var newWidth = (int)(widthInPixels * Scale) + lMarginPixel + rMarginPixel;
            new ShapeSize(parent).UpdateWidth(newWidth);
        }
    }
    
    private void UpdateHeight(
        float textWidth,
        int currentBlockWidth,
        float textHeight,
        int tMarginPixel,
        int bMarginPixel,
        int currentBlockHeight,
        OpenXmlElement parent)
    {
        var requiredRowsCount = textWidth / currentBlockWidth;
        var integerPart = (int)requiredRowsCount;
        var fractionalPart = requiredRowsCount - integerPart;
        if (fractionalPart > 0)
        {
            integerPart++;
        }

        var requiredHeight = (integerPart * textHeight) + tMarginPixel + bMarginPixel;
        var newHeight = (int)requiredHeight + tMarginPixel + bMarginPixel + tMarginPixel + bMarginPixel;
        var position = new Position(parent);
        var size = new ShapeSize(parent);
        size.UpdateHeight(newHeight);

        // We should raise the shape up by the amount which is half of the increased offset.
        // PowerPoint does the same thing.
        var yOffset = (requiredHeight - currentBlockHeight) / 2;
        position.UpdateY((int)(position.Y() - yOffset));
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