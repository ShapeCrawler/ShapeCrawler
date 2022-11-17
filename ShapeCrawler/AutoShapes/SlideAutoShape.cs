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
    private readonly Lazy<ShapeFill> shapeFill;
    private readonly Lazy<TextFrame?> textFrame;
    private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
    private readonly P.Shape pShape;

    internal SlideAutoShape(P.Shape pShape, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide,
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

        var newTextRect = default(SKRect);
        var newText = this.TextFrame.Text;
        paint.MeasureText(newText, ref newTextRect);
        var newTextW = newTextRect.Width;
        var newTextH = paint.TextSize;
        var textBlockW = this.Width - lMarginPixel - rMarginPixel;
        var currentTextBlockHeight = this.Height - tMarginPixel - bMarginPixel;

        var hasTextBlockEnoughWidth = textBlockW > newTextW;
        if (hasTextBlockEnoughWidth && this.TextFrame.TextWrapped)
        {
            return;
        }

        var neededRowsCount = newTextW / textBlockW;
        var integerPart = (int)neededRowsCount;
        var fractionalPart = neededRowsCount - integerPart;
        if (fractionalPart > 0)
        {
            integerPart++;
        }

        var requiredTextBlockHeight = (integerPart * newTextH) + tMarginPixel + bMarginPixel;
        var hasRequiredHeight = currentTextBlockHeight >= requiredTextBlockHeight;
        if (hasRequiredHeight && this.TextFrame.TextWrapped)
        {
            return;
        }

        this.Height = (int)requiredTextBlockHeight + tMarginPixel + bMarginPixel + tMarginPixel + bMarginPixel;

        // We should raise the shape up by the amount which is half of the increased offset.
        // PowerPoint does the same thing.
        var yOffset = (requiredTextBlockHeight - currentTextBlockHeight) / 2;
        this.Y -= (int)yOffset;

        if (!this.TextFrame.TextWrapped)
        {
            var longerText = this.TextFrame.Paragraphs
                .Select(x => new { x.Text, x.Text.Length })
                .OrderByDescending(x => x.Length)
                .First().Text;
            var paraTextRect = default(SKRect);
            var widthInPixels = paint.MeasureText(longerText, ref paraTextRect);
            this.Width = (int)(widthInPixels * 1.4) + lMarginPixel + rMarginPixel;
        }
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