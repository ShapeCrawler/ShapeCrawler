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
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler;

/// <summary>
///     Represents AutoShape located on Slide.
/// </summary>
internal class SlideAutoShape : SlideShape, IAutoShape, ITextFrameContainer
{
    private readonly Lazy<ShapeFill> shapeFill;
    private readonly Lazy<TextFrame?> textFrame;
    private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
    private readonly P.Shape pShape;

    internal SlideAutoShape(P.Shape pShape, OneOf<SCSlide, SCSlideLayout, SCSlideMaster> oneOfSlide, SCGroupShape groupShape)
        : base(pShape, oneOfSlide, groupShape)
    {
        this.pShape = pShape;
        this.textFrame = new Lazy<TextFrame?>(this.GetTextBox);
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

            var textFrameInternal = (TextFrame)this.TextFrame;
            if (textFrameInternal != null)
            {
                textFrameInternal.Draw(canvas, rect);
            }
        }
    }

    internal void FillFontData(int paragraphLvl, ref FontData fontData)
    {
        if (this.lvlToFontData.Value.TryGetValue(paragraphLvl, out FontData layoutFontData))
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

    private TextFrame? GetTextBox()
    {
        var pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
        return pTextBody == null ? null : new TextFrame(this, pTextBody);
    }

    private ShapeFill GetFill() // TODO: duplicate of LayoutAutoShape.TryGetFill()
    {
        return new ShapeFill(this);
    }
}