using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal abstract class CopyableShape : Shape
{
    internal CopyableShape(OpenXmlElement openXmlElement)
        : base(openXmlElement)
    {
    }

    internal virtual void CopyTo(P.ShapeTree pShapeTree)
    {
        new SCPShapeTree(pShapeTree).Add(this.PShapeTreeElement);
    }
}