using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using OneOf;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Drawing.ShapeFill;
using ShapeCrawler.Extensions;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.Texts;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents interface of AutoShape.
/// </summary>
public interface IAutoShape : IShape
{
    /// <summary>
    ///     Gets shape outline.
    /// </summary>
    IShapeOutline Outline { get; }
 
    /// <summary>
    ///     Gets shape fill. Returns <see langword="null"/> if the shape can not be filled.
    /// </summary>
    IShapeFill? Fill { get; }
    
    /// <summary>
    ///     Gets text frame. Returns <see langword="null"/> if the shape is not a text holder.
    /// </summary>
    ITextFrame? TextFrame { get; }
    
    /// <summary>
    ///     Duplicate the shape.
    /// </summary>
    IAutoShape Duplicate();
}

internal class SCAutoShape : SCShape, IAutoShape, ITextFrameContainer
{
    // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi.
    // 96/72=1.4
    private const double Scale = 1.4;

    private readonly Lazy<SCShapeFill> shapeFill;
    private readonly Lazy<SCTextFrame?> textFrame;
    private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
    private readonly TypedOpenXmlCompositeElement pShape;

    internal SCAutoShape(
        TypedOpenXmlCompositeElement pShape,
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> parentSlideStructureOf,
        OneOf<ShapeCollection, SCGroupShape> parentShapeCollection)
        : base(pShape, parentSlideStructureOf, parentShapeCollection)
    {
        this.pShape = pShape;
        this.textFrame = new Lazy<SCTextFrame?>(this.GetTextFrame);
        this.shapeFill = new Lazy<SCShapeFill>(this.GetFill);
        this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
    }
    
    internal event EventHandler<NewAutoShape>? Duplicated;

    #region Public Properties

    public IShapeOutline Outline => this.GetOutline();

    public SCShape SCShape => this; // TODO: should be internal?

    public override SCShapeType ShapeType => SCShapeType.AutoShape;

    public virtual IShapeFill? Fill => this.shapeFill.Value;
    
    public virtual ITextFrame? TextFrame => this.textFrame.Value;

    public virtual IAutoShape Duplicate()
    {
        var typedCompositeElement = (TypedOpenXmlCompositeElement)this.PShapeTreeChild.CloneNode(true);
        var id = ((SlideStructure)this.ParentSlideStructureOf.Value).GetNextShapeId();
        typedCompositeElement.GetNonVisualDrawingProperties().Id = new UInt32Value((uint)id);
        var newAutoShape = new SCAutoShape(
            typedCompositeElement, 
            this.ParentSlideStructureOf,
            this.ParentShapeCollectionStructureOf);

        var duplicatedShape = new NewAutoShape(newAutoShape, typedCompositeElement);
        this.Duplicated?.Invoke(this, duplicatedShape);
        
        return newAutoShape;
    }

    #endregion Public Properties

    internal override void Draw(SKCanvas slideCanvas)
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

    internal void ResizeShape()
    {
        if (this.TextFrame!.AutofitType != SCAutofitType.Resize)
        {
            return;
        }

        var baseParagraph = this.TextFrame.Paragraphs.First();
        var popularPortion = baseParagraph.Portions.GroupBy(p => p.Font.Size).OrderByDescending(x => x.Count())
            .First().First();
        var font = popularPortion.Font;

        var paint = new SKPaint();
        var fontSize = font.Size;
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
                var referencedMasterShape = (SCAutoShape?)placeholder.ReferencedShape.Value;
                referencedMasterShape?.FillFontData(paragraphLvl, ref fontData);
            }

            return;
        }

        if (this.Placeholder != null)
        {
            var placeholder = (SCPlaceholder)this.Placeholder;
            var referencedMasterShape = (SCAutoShape?)placeholder.ReferencedShape.Value;
            if (referencedMasterShape != null)
            {
                referencedMasterShape.FillFontData(paragraphLvl, ref fontData);
            }
        }
    }

    private Dictionary<int, FontData> GetLvlToFontData()
    {
        var textBody = this.pShape.GetFirstChild<P.TextBody>();
        var lvlToFontData = FontDataParser.FromCompositeElement(textBody!.ListStyle!);

        if (!lvlToFontData.Any())
        {
            var endParaRunPrFs = textBody.GetFirstChild<A.Paragraph>() !
                .GetFirstChild<A.EndParagraphRunProperties>()?.FontSize;
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

    private SCTextFrame? GetTextFrame()
    {
        var pTextBody = this.PShapeTreeChild.GetFirstChild<P.TextBody>();
        if (pTextBody == null)
        {
            return null;
        }

        var newTextFrame = new SCTextFrame(this, pTextBody);
        newTextFrame.TextChanged += this.ResizeShape;

        return newTextFrame;
    }

    private SCShapeFill GetFill()
    {
        var slideObject = (SlideStructure)this.SlideStructure;
        return new SCAutoShapeFill(slideObject, this.pShape.GetFirstChild<P.ShapeProperties>() !, this);
    }

    private IShapeOutline GetOutline()
    {
        return new SCShapeOutline(this);
    }
}