using DocumentFormat.OpenXml;
using ShapeCrawler.Drawing;
using ShapeCrawler.Slides;
using Position = ShapeCrawler.Positions.Position;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class MediaShape : Shape
{
    internal MediaShape(Position position, ShapeSize shapeSize, ShapeId shapeId, P.Picture pPicture):
        base(position, shapeSize, shapeId, pPicture)
    {
        this.Media = new Media(new SlideShapeOutline(pPicture.ShapeProperties!), new ShapeFill(pPicture.ShapeProperties!), pPicture);
    }

    public override IMedia? Media { get; }
}