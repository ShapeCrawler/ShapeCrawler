using DocumentFormat.OpenXml;
using ShapeCrawler.Shapes;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.SmartArts;

internal sealed class SmartArtShape : DrawingShape
{
    internal SmartArtShape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement)
        : base(position, shapeSize, shapeId, pShapeTreeElement)
    {
        this.SmartArt = new SmartArt(new SmartArtNodeCollection());
    }

    public override ISmartArt? SmartArt { get; }
}