using System;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Drawing;
using ShapeCrawler.Services;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler.AutoShapes;

internal class LayoutAutoShape : LayoutShape, IAutoShape, ITextFrameContainer
{
    private readonly ResettableLazy<Dictionary<int, FontData>> lvlToFontData;
    private readonly Lazy<ShapeFill> shapeFill;
    private readonly Lazy<TextFrame?> textBox;
    private readonly P.Shape pShape;

    internal LayoutAutoShape(SCSlideLayout slideLayout, P.Shape pShape)
        : base(slideLayout, pShape)
    {
        this.textBox = new Lazy<TextFrame?>(this.GetTextFrame);
        this.shapeFill = new Lazy<ShapeFill>(TryGetFill);

        this.pShape = pShape;
    }

    public ITextFrame? TextFrame => this.textBox.Value;

    public IShapeFill Fill => this.shapeFill.Value;

    public override SCShapeType ShapeType => SCShapeType.AutoShape;

    public Shape Shape => this;

    private TextFrame? GetTextFrame()
    {
        P.TextBody pTextBody = this.PShapeTreesChild.GetFirstChild<P.TextBody>();
        if (pTextBody == null)
        {
            return null;
        }

        IEnumerable<A.Text> aTexts = pTextBody.Descendants<A.Text>();
        if (aTexts.Sum(t => t.Text.Length) > 0)
        {
            return new TextFrame(this, pTextBody);
        }

        return null;
    }

    private static ShapeFill TryGetFill() // TODO: duplicate of SlideAutoShape.TryGetFill()
    {
        throw new NotImplementedException();
    }
}