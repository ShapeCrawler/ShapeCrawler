using System;
using System.Collections.Generic;
using System.Linq;
using AngleSharp.Html.Dom;
using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed record SCSlideAutoShape : IAutoShape
{
    // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi.
    // 96/72=1.4
    private const double Scale = 1.4;

    private readonly Lazy<SCShapeFill> shapeFill;
    private readonly Lazy<SCTextFrame?> textFrame;
    private readonly ResetableLazy<Dictionary<int, FontData>> lvlToFontData;
    private readonly P.Shape pShape;
    private readonly IReadOnlyShapeCollection parentShapeCollection;
    private readonly Shape shape;

    internal SCSlideAutoShape(
        P.Shape pShape,
        IReadOnlyShapeCollection parentShapeCollection)
    {
        this.pShape = pShape;
        this.parentShapeCollection = parentShapeCollection;
        this.textFrame = new Lazy<SCTextFrame?>(this.ParseTextFrame);
        this.shapeFill = new Lazy<SCShapeFill>(this.GetFill);
        this.lvlToFontData = new ResetableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
        this.shape = new Shape(pShape);
        this.Outline = new SCShapeOutline(this, pShape.ShapeProperties!);
    }
    
    internal event EventHandler<NewAutoShape>? Duplicated;

    #region Public Properties

    public IShapeOutline Outline { get; }

    public int Width
    {
        get => this.shape.Width(); 
        set => this.shape.UpdateWidth(value);
    }

    public int Height
    {
        get => this.shape.Height(); 
        set => this.shape.UpdateHeight(value);
    }
    public int Id => this.shape.Id();
    public string Name => this.shape.Name();
    public bool Hidden { get; }
    public IPlaceholder? Placeholder { get; }
    public SCGeometry GeometryType { get; }
    public string? CustomData { get; set; }
    public SCShapeType ShapeType => SCShapeType.AutoShape;
    public IAutoShape AsAutoShape()
    {
        return this;
    }

    public IShapeFill Fill => this.shapeFill.Value;

    public ITextFrame? TextFrame => this.textFrame.Value;

    public void Duplicate()
    {
        this.parentShapeCollection.Copy(this.Id);
        var typedCompositeElement = (TypedOpenXmlCompositeElement)this.pShape.CloneNode(true);
        var id = this.GetNextShapeId();
        typedCompositeElement.GetNonVisualDrawingProperties().Id = new UInt32Value((uint)id);
        var newAutoShape = new SCSlideAutoShape((P.Shape)typedCompositeElement, this.parentShapeCollection);

        var duplicatedShape = new NewAutoShape(newAutoShape, typedCompositeElement);
        this.Duplicated?.Invoke(this, duplicatedShape);
    }

    private int GetNextShapeId()
    {
        if (this.parentShapeCollection.Any())
        {
            return slide.Shapes.Select(shape => shape.Id).Prepend(0).Max() + 1;    
        }

        return 1;
    }
    
    #endregion Public Properties

    internal void Draw(SKCanvas slideCanvas)
    {
        var skColorOutline = SKColor.Parse(this.Outline.Color);

        using var paint = new SKPaint
        {
            Color = skColorOutline,
            IsAntialias = true,
            StrokeWidth = UnitConverter.PointToPixel(this.Outline.Weight),
            Style = SKPaintStyle.Stroke
        };

        if (this.GeometryType == SCGeometry.Rectangle)
        {
            float left = this.X;
            float top = this.Y;
            float right = this.X + this.Width;
            float bottom = this.Y + this.Height;
            var rect = new SKRect(left, top, right, bottom);
            slideCanvas.DrawRect(rect, paint);
            var textFrame = (SCTextFrame)this.TextFrame!;
            textFrame.Draw(slideCanvas, left, this.Y);
        }
    }

    internal string ToJson()
    {
        throw new NotImplementedException();
    }

    internal IHtmlElement ToHtmlElement()
    {
        throw new NotImplementedException();
    }

    internal void ResizeShape()
    {
        if (this.TextFrame!.AutofitType != SCAutofitType.Resize)
        {
            return;
        }

        var baseParagraph = this.TextFrame.Paragraphs.First();
        var popularPortion = baseParagraph.Portions.OfType<SCRegularPortion>().GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count())
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

    internal void FillFontData(int paragraphLvl, ref FontData fontData)
    {
        if (this.lvlToFontData.Value.TryGetValue(paragraphLvl, out var layoutFontData))
        {
            fontData = layoutFontData;
            if (!fontData.IsFilled() && this.Placeholder != null)
            {
                var placeholder = (SCPlaceholder)this.Placeholder;
                var referencedMasterShape = (SCSlideAutoShape?)placeholder.ReferencedShape.Value;
                referencedMasterShape?.FillFontData(paragraphLvl, ref fontData);
            }

            return;
        }

        if (this.Placeholder != null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder;
            var referencedMasterShape = (SCSlideAutoShape?)placeholder.ReferencedShape.Value;
            if (referencedMasterShape != null)
            {
                referencedMasterShape.FillFontData(paragraphLvl, ref fontData);
            }
        }
    }

    private Dictionary<int, FontData> GetLvlToFontData()
    {
        var textBody = this.pShape.GetFirstChild<DocumentFormat.OpenXml.Presentation.TextBody>();
        var lvlToFontData = FontDataParser.FromCompositeElement(textBody!.ListStyle!);

        if (!lvlToFontData.Any())
        {
            var endParaRunPrFs = textBody.GetFirstChild<DocumentFormat.OpenXml.Drawing.Paragraph>() !
                .GetFirstChild<DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties>()?.FontSize;
            if (endParaRunPrFs is not null)
            {
                var fontData = new FontData
                {
                    FontSize = endParaRunPrFs
                };
                lvlToFontData.Add(1, fontData);
            }
        }

        return lvlToFontData;
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

    private SCTextFrame? ParseTextFrame()
    {
        var pTextBody = this.PShapeTreeChild.GetFirstChild<DocumentFormat.OpenXml.Presentation.TextBody>();
        if (pTextBody == null)
        {
            return null;
        }

        var newTextFrame = new SCTextFrame(this, pTextBody, this.slideStructure, this);
        newTextFrame.TextChanged += this.ResizeShape;

        return newTextFrame;
    }

    private SCShapeFill GetFill()
    {
        var slideObject = this.SlideStructure;
        return new SCAutoShapeFill(
            slideObject, 
            this.pShape.GetFirstChild<P.ShapeProperties>() !, 
            this, 
            this.sdkSlidePart);
    }
    
    public int X { get; set; }
    public int Y { get; set; }
}