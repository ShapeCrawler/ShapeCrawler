using System;
using System.Collections.Generic;
using System.Linq;
using OneOf;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Factories;
using ShapeCrawler.Placeholders;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Statics;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler;

internal class SlideAutoShape : SlideShape, IAutoShape, ITextFrameContainer
{
    // SkiaSharp uses 72 Dpi (https://stackoverflow.com/a/69916569/2948684), ShapeCrawler uses 96 Dpi.
    // 96/72=1.4
    private const double Scale = 1.4;

    private readonly Lazy<ShapeFill> shapeFill;
    private readonly Lazy<TextFrame?> textFrame;
    private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
    private readonly P.Shape pShape;

    internal SlideAutoShape(
        P.Shape pShape, 
        OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide,
        SCGroupShape groupShape)
        : base(pShape, oneOfSlide, groupShape)
    {
        this.pShape = pShape;
        this.textFrame = new Lazy<TextFrame?>(this.GetTextFrame);
        this.shapeFill = new Lazy<ShapeFill>(this.GetFill);
        this.lvlToFontData = new ResettableLazy<Dictionary<int, FontData>>(this.GetLvlToFontData);
    }

    #region Public Properties

    public IShapeFill Fill => this.shapeFill.Value;

    public Shape Shape => this; // TODO: should be internal?

    public override SCShapeType ShapeType => SCShapeType.AutoShape;

    public ITextFrame? TextFrame => this.textFrame.Value;

    #endregion Public Properties

    internal override void Draw(SKCanvas canvas)
    {
        var paint = new SKPaint
        {
            Color = SKColors.Black,
            IsAntialias = true,
            Style = SKPaintStyle.Fill,
            TextAlign = SKTextAlign.Center,
            TextSize = 12
        };

        if (this.GeometryType == SCGeometry.Rectangle)
        {
            var rect = new SKRect(this.X, this.Y, this.Width, this.Height);
            canvas.DrawRect(rect, paint);

            var textFrameInternal = this.TextFrame as TextFrame;
            if (textFrameInternal != null)
            {
                textFrameInternal.Draw(canvas, rect);
            }
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
        paint.Typeface = SKTypeface.FromFamilyName(font.Name);
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
                var placeholder = (Placeholder)this.Placeholder;
                var referencedMasterShape = (SlideAutoShape)placeholder.ReferencedShape;
                referencedMasterShape?.FillFontData(paragraphLvl, ref fontData);
            }

            return;
        }

        if (this.Placeholder != null)
        {
            var placeholder = (Placeholder)this.Placeholder;
            var referencedMasterShape = (SlideAutoShape)placeholder.ReferencedShape;
            if (referencedMasterShape != null)
            {
                referencedMasterShape.FillFontData(paragraphLvl, ref fontData);
            }
        }
    }

    private Dictionary<int, FontData> GetLvlToFontData()
    {
        var textBody = this.pShape.TextBody!;
        var lvlToFontData = FontDataParser.FromCompositeElement(textBody.ListStyle!);

        if (!lvlToFontData.Any())
        {
            var endParaRunPrFs = textBody.GetFirstChild<A.Paragraph>()!
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

    private TextFrame? GetTextFrame()
    {
        var pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
        if (pTextBody == null)
        {
            return null;
        }

        var newTextFrame = new TextFrame(this, pTextBody);
        newTextFrame.TextChanged += this.ResizeShape;

        return newTextFrame;
    }

    private ShapeFill GetFill()
    {
        return new ShapeFill(this);
    }
}