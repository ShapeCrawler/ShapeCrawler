using System.Linq;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class OleObject : Shape
{
    private readonly P.GraphicFrame pGraphicFrame;

    internal OleObject(P.GraphicFrame pGraphicFrame)
        : base(pGraphicFrame)
    {
        this.pGraphicFrame = pGraphicFrame;
        this.Outline = new SlideShapeOutline(pGraphicFrame.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(pGraphicFrame.Descendants<P.ShapeProperties>().First());
    }

    public override ShapeContent ShapeContent => ShapeContent.OleObject;

    public override bool HasOutline => true;

    public override IShapeOutline Outline { get; }

    public override bool HasFill => true;

    public override IShapeFill Fill { get; }

    public override bool Removeable => true;

    public override void Remove() => this.pGraphicFrame.Remove();
}