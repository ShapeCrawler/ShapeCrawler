using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class OleObject : Shape
{
    private readonly P.GraphicFrame pGraphicFrame;

    internal OleObject(OpenXmlPart openXmlPart, P.GraphicFrame pGraphicFrame)
        : base(openXmlPart, pGraphicFrame)
    {
        this.pGraphicFrame = pGraphicFrame;
        this.Outline = new SlideShapeOutline(openXmlPart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(openXmlPart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
    }

    public override ShapeType ShapeType => ShapeType.OleObject;

    public override bool HasOutline => true;

    public override IShapeOutline Outline { get; }

    public override bool HasFill => true;

    public override IShapeFill Fill { get; }

    public override bool Removeable => true;

    public override void Remove() => this.pGraphicFrame.Remove();
}