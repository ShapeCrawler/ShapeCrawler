using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeSize(OpenXmlPart openXmlPart, OpenXmlElement pShapeTreeElement)
{
    internal decimal Width
    {
        get => new Emus(this.GetAExtents().Cx!).AsPoints();
        set => this.GetAExtents().Cx = new Points(value).AsEmus();
    }

    internal decimal Height
    {
        get => new Emus(this.GetAExtents().Cy!).AsPoints();
        set => this.GetAExtents().Cy = new Points(value).AsEmus();
    }

    private A.Extents GetAExtents()
    {
        var aExtents = pShapeTreeElement.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents != null)
        {
            return aExtents;
        }

        return new ReferencedPShape(openXmlPart, pShapeTreeElement).ATransform2D().Extents!;
    }
}