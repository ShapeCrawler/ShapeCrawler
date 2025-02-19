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

    internal decimal X() => new Emus(this.AOffset().X!.Value).AsHorizontalPixels();
   
    internal void UpdateX(decimal pixels)
    {
        var emus = new Pixels(pixels).AsHorizontalEmus();
        this.AOffset().X = new Int64Value(emus);
    }
    
    internal decimal Y() => new Emus(this.AOffset().Y!.Value).AsVerticalPixels();

    internal void UpdateY(decimal pixels)
    {
        var emus = new Pixels(pixels).AsVerticalEmus();
        this.AOffset().Y = new Int64Value(emus);
    }

    private A.Offset AOffset()
    {
        var aOffset = this.pShapeTreeElement.Descendants<A.Offset>().FirstOrDefault();
        if (aOffset != null)
        {
            return aOffset;
        }

        return new ReferencedPShape(this.sdkTypedOpenXmlPart, this.pShapeTreeElement).ATransform2D().Offset!;
    }  
}