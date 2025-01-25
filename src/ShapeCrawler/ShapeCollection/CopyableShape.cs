using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal abstract class CopyableShape : Shape
{
    internal CopyableShape(OpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement openXmlElement)
        : base(sdkTypedOpenXmlPart, openXmlElement)
    {
    }

    internal virtual void CopyTo(
        int id,
        P.ShapeTree pShapeTree,
        IEnumerable<string> existingShapeNames)
    {
        new SPShapeTree(pShapeTree).Add(this.PShapeTreeElement);
    }
}