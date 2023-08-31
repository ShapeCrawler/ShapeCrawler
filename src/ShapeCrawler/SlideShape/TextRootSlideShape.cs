using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.SlideShape;

/// <summary>
///     A text AutoShape on a slide.
/// </summary>
internal sealed record TextRootSlideShape : IRootSlideShape
{
    private readonly SlidePart sdkSlidePart;
    private readonly IRootSlideShape rootSlideShape;
    private readonly P.TextBody sdkPTextBody;
    private readonly Lazy<TextFrame> textFrame;

    // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi.
    // 96/72=1.4
    private const double Scale = 1.4;

    internal TextRootSlideShape(SlidePart sdkSlidePart, IRootSlideShape rootSlideShape, P.TextBody sdkPTextBody)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.rootSlideShape = rootSlideShape;
        this.sdkPTextBody = sdkPTextBody;
        this.textFrame = new Lazy<TextFrame>(this.ParseTextFrame);
    }

    private TextFrame ParseTextFrame()
    {
        var newTextFrame = new TextFrame(this.sdkSlidePart, this.sdkPTextBody);
        newTextFrame.TextChanged += this.ResizeShape;

        return newTextFrame;
    }

    private void ResizeShape()
    {
        if (this.TextFrame!.AutofitType != SCAutofitType.Resize)
        {
            return;
        }

        var baseParagraph = this.TextFrame.Paragraphs.First();
        var popularPortion = baseParagraph.Portions.OfType<TextParagraphPortion>().GroupBy(p => p.Font.Size)
            .OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var paint = new SKPaint();
        var fontSize = font!.Size;
        paint.TextSize = fontSize;
        paint.Typeface = SKTypeface.FromFamilyName(font.LatinName);
        paint.IsAntialias = true;

        var lMarginPixel = UnitConverter.CentimeterToPixel(this.TextFrame.LeftMargin);
        var rMarginPixel = UnitConverter.CentimeterToPixel(this.TextFrame.RightMargin);
        var tMarginPixel = UnitConverter.CentimeterToPixel(this.TextFrame.TopMargin);
        var bMarginPixel = UnitConverter.CentimeterToPixel(this.TextFrame.BottomMargin);

        var textRect = default(SKRect);
        var text = this.TextFrame.Text;
        paint.MeasureText(text, ref textRect);
        var textWidth = textRect.Width;
        var textHeight = paint.TextSize;
        var currentBlockWidth = this.Width - lMarginPixel - rMarginPixel;
        var currentBlockHeight = this.Height - tMarginPixel - bMarginPixel;

        this.UpdateHeight(textWidth, currentBlockWidth, textHeight, tMarginPixel, bMarginPixel, currentBlockHeight);
        this.UpdateWidthIfNeed(paint, lMarginPixel, rMarginPixel);
    }
    
    private void UpdateHeight(
        float textWidth,
        int currentBlockWidth,
        float textHeight,
        int tMarginPixel,
        int bMarginPixel,
        int currentBlockHeight)
    {
        var requiredRowsCount = textWidth / currentBlockWidth;
        var integerPart = (int)requiredRowsCount;
        var fractionalPart = requiredRowsCount - integerPart;
        if (fractionalPart > 0)
        {
            integerPart++;
        }

        var requiredHeight = (integerPart * textHeight) + tMarginPixel + bMarginPixel;
        this.Height = (int)requiredHeight + tMarginPixel + bMarginPixel + tMarginPixel + bMarginPixel;

        // We should raise the shape up by the amount which is half of the increased offset.
        // PowerPoint does the same thing.
        var yOffset = (requiredHeight - currentBlockHeight) / 2;
        this.Y -= (int)yOffset;
    }

    private void UpdateWidthIfNeed(SKPaint paint, int lMarginPixel, int rMarginPixel)
    {
        if (!this.TextFrame!.TextWrapped)
        {
            var longerText = this.TextFrame.Paragraphs
                .Select(x => new { x.Text, x.Text.Length })
                .OrderByDescending(x => x.Length)
                .First().Text;
            var paraTextRect = default(SKRect);
            var widthInPixels = paint.MeasureText(longerText, ref paraTextRect);
            this.Width = (int)(widthInPixels * Scale) + lMarginPixel + rMarginPixel;
        }
    }

    public bool IsTextHolder => true;

    public ITextFrame TextFrame => this.textFrame.Value;
    
    #region Slide Properties
    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.rootSlideShape.Id;
    public string Name => this.rootSlideShape.Name;
    public bool Hidden => this.Hidden;

    public bool IsPlaceholder => false;

    public IPlaceholder Placeholder => throw new SCException(
        $"Text shape is not a placeholder. Use {nameof(IShape.IsPlaceholder)} property to check if the shape is a placeholder.");
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }
    public bool HasOutline { get; }
    public IShapeOutline Outline => this.rootSlideShape.Outline;
    public IShapeFill Fill => this.rootSlideShape.Fill;
    public double Rotation { get; }
    public void Duplicate() => this.rootSlideShape.Duplicate();
    #endregion Slide Properties
}