using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Positions;

internal sealed class Position
{
    private readonly OpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlElement pShapeTreeElement;

    internal Position(OpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.pShapeTreeElement = pShapeTreeElement;
    }

    internal decimal X
    {
        get
        {
            var emus = this.GetAOffset().X!.Value;
            return new Emus(emus).AsPoints();
        }
        set
        {
            var emus = new Points(value).AsEmus();
            this.GetAOffset().X = new Int64Value(emus);    
        }
    }
    
    internal decimal Y
    {
        get
        {
            return new Emus(this.GetAOffset().Y!.Value).AsPoints();
        }
        set
        {
            var emus = new Points(value).AsEmus();
            this.GetAOffset().Y = new Int64Value(emus);
        }
    }

    private A.Offset GetAOffset()
    {
        var aOffset = this.pShapeTreeElement.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset != null)
        {
            return aOffset;
        }

        return new ReferencedPShape(this.sdkTypedOpenXmlPart, this.pShapeTreeElement).ATransform2D().Offset!;
    }  
}