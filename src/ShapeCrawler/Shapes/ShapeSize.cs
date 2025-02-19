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

    internal decimal Height() => UnitConverter.VerticalEmuToPixel(this.AExtents().Cy!);

    internal void UpdateHeight(decimal heightPixels) => this.AExtents().Cy = UnitConverter.VerticalPixelToEmu(heightPixels);

    internal decimal Width() => UnitConverter.HorizontalEmuToPixel(this.AExtents().Cx!);

    internal void UpdateWidth(decimal widthPixels) => this.AExtents().Cx = UnitConverter.HorizontalPixelToEmu(widthPixels);

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