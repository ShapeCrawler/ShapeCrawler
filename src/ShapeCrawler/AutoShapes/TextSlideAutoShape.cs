using System;
using System.Linq;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed record TextSlideAutoShape : ISlideAutoShape
{
    private readonly SlideAutoShape autoShape;
    private readonly P.TextBody pTextBody;
    private readonly Lazy<TextFrame> textFrame;
    
    // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi.
    // 96/72=1.4
    private const double Scale = 1.4;

    internal TextSlideAutoShape(SlideAutoShape autoShape, P.TextBody pTextBody)
    {
        this.autoShape = autoShape;
        this.pTextBody = pTextBody;
        this.textFrame = new Lazy<TextFrame>(this.ParseTextFrame);
    }

    private TextFrame ParseTextFrame()
    {
        var newTextFrame = new TextFrame(this, pTextBody);
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
        var popularPortion = baseParagraph.Portions.OfType<SCParagraphTextPortion>().GroupBy(p => p.Font.Size)
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

    public ITextFrame TextFrame => this.textFrame.Value;
    
    #region AutoShape Properties

    public int X { get; set; }
    public int Y { get; set; }
    public int Width { get; set; }
    public int Height { get; set; }
    public int Id => this.autoShape.Id;
    public string Name => this.autoShape.Name;
    public bool Hidden => this.Hidden;

    public bool IsPlaceholder() => this.autoShape.IsPlaceholder();

    public IPlaceholder Placeholder => this.autoShape.Placeholder;
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType { get; }

    public IAutoShape AsAutoShape() => this;

    public IShapeOutline Outline => this.autoShape.Outline;
    public IShapeFill Fill => this.autoShape.Fill;
    public bool IsTextHolder() => true;

    public double Rotation { get; }
    public void Duplicate() => this.autoShape.Duplicate();

    #endregion AutoShape Properties
}