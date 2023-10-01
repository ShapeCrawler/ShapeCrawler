using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Shapes;

internal record struct ReferencedPShape
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlElement pShapeTreeElement;

    internal ReferencedPShape(TypedOpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.pShapeTreeElement = pShapeTreeElement;
    }

    internal Transform2D ATransform2D()
    {
        throw new NotImplementedException();
    }
}