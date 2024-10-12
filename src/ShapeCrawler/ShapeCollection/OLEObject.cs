using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal sealed class OleObject : Shape
{
    private readonly P.GraphicFrame pGraphicFrame;

    internal OleObject(OpenXmlPart sdkTypedOpenXmlPart, P.GraphicFrame pGraphicFrame)
        : base(sdkTypedOpenXmlPart, pGraphicFrame)
    {
        this.pGraphicFrame = pGraphicFrame;
        this.Outline = new SlideShapeOutline(sdkTypedOpenXmlPart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
        this.Fill = new ShapeFill(sdkTypedOpenXmlPart, pGraphicFrame.Descendants<P.ShapeProperties>().First());
    }

    public override ShapeType ShapeType => ShapeType.OleObject;

    public override bool HasOutline => true;
    
    public override IShapeOutline Outline { get; }
    
    public override bool HasFill => true;
    
    public override IShapeFill Fill { get; }
    
    public override bool Removeable => true;
    
    public override void Remove() => this.pGraphicFrame.Remove();
}