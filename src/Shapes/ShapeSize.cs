using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed class ShapeSize(OpenXmlElement pShapeTreeElement)
{
    internal decimal Width
    {
        get => new Emus(this.GetAExtents().Cx!).AsPoints();
        set
        {
            var emus = new Points(value).AsEmus();
            if (emus > int.MaxValue)
            {
                emus = int.MaxValue;
            }

            this.GetAExtents().Cx = emus;
        }
    }

    internal decimal Height
    {
        get => new Emus(this.GetAExtents().Cy!).AsPoints();
        set
        {
            var emus = new Points(value).AsEmus();
            if (emus > int.MaxValue)
            {
                emus = int.MaxValue;
            }

            this.GetAExtents().Cy = emus;
        }
    }

    private A.Extents GetAExtents()
    {
        var aExtents = pShapeTreeElement.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents != null)
        {
            return aExtents;
        }

        return new ReferencedPShape(pShapeTreeElement).ATransform2D().Extents!;
    }
}