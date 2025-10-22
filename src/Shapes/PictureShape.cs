using ShapeCrawler.Positions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal class PictureShape(Picture picture, P.Picture pPicture) : Shape(new Position(pPicture), new ShapeSize(pPicture), new ShapeId(pPicture), pPicture)
{
    public override IPicture Picture => picture;

    public override void CopyTo(P.ShapeTree pShapeTree)
    {
        picture.CopyTo(pShapeTree);
    }
}