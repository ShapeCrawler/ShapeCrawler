using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

internal class PictureShape(Picture picture, P.Picture pPicture) : Shape(new Position(pPicture), new ShapeSize(pPicture), new ShapeId(pPicture), pPicture)
{
    public override IPicture Picture => picture;

    public override void CopyTo(P.ShapeTree pShapeTree)
    {
        picture.CopyTo(pShapeTree);
    }
}