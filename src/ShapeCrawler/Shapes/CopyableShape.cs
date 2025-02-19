using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal abstract class CopyableShape : Shape
{
    internal CopyableShape(OpenXmlPart openXmlPart, OpenXmlElement openXmlElement)
        : base(openXmlPart, openXmlElement)
    {
    }

    internal virtual void CopyTo(P.ShapeTree pShapeTree)
    {
        new SCPShapeTree(pShapeTree).Add(this.PShapeTreeElement);
    }
}