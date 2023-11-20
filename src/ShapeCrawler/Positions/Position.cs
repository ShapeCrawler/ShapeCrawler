using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Positions;

internal sealed class Position
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlElement pShapeTreeElement;

    internal Position(TypedOpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.pShapeTreeElement = pShapeTreeElement;
    }

    internal int X() => new Emus(this.AOffset().X!.Value).AsHorizontalPixels();
   
    internal void UpdateX(int pixels)
    {
        var emus = new Pixels(pixels).AsHorizontalEmus();
        this.AOffset().X = new Int64Value(emus);
    }
    
    internal int Y() => new Emus(this.AOffset().Y!.Value).AsVerticalPixels();

    internal void UpdateY(int pixels)
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