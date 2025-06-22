using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Groups;

internal class PictureShape(Picture picture, P.Picture pPicture) : Shape(new Position(pPicture), new ShapeSize(pPicture), new ShapeId(pPicture), pPicture)
{
    public override Picture Picture => picture;
}