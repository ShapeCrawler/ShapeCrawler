using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeSize
{
    private readonly OpenXmlPart openXmlPart;
    private readonly OpenXmlElement pShapeTreeElement;

    internal ShapeSize(OpenXmlPart openXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.openXmlPart = openXmlPart;
        this.pShapeTreeElement = pShapeTreeElement;
    }

    internal float Width
    {
        get
        {
            return new Emus(this.AExtents().Cx!).AsPoints();
        }
        set
        {
            var emus = new Points(value).AsEmus();
            this.AExtents().Cx = emus;
        }
    }
    
    internal float Height
    {
        get
        {
            return new Emus(this.AExtents().Cy!).AsPoints();
        }
        set
        {
            var emus = new Points(value).AsEmus();
            this.AExtents().Cy = emus;
        }
    }
    
    private A.Extents AExtents()
    {
        var aExtents = this.pShapeTreeElement.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents != null)
        {
            return aExtents;
        }

        return new ReferencedPShape(this.openXmlPart, this.pShapeTreeElement).ATransform2D().Extents!;
    }
}