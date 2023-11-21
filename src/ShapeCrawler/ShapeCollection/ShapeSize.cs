using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.ShapeCollection;

internal sealed class ShapeSize
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlElement sdkPShapeTreeElement;

    internal ShapeSize(TypedOpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement sdkPShapeTreeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.sdkPShapeTreeElement = sdkPShapeTreeElement;
    }

    internal int Height() => UnitConverter.VerticalEmuToPixel(this.AExtents().Cy!);
   
    internal void UpdateHeight(int heightPixels) => this.AExtents().Cy = UnitConverter.VerticalPixelToEmu(heightPixels);
    
    internal int Width() => UnitConverter.HorizontalEmuToPixel(this.AExtents().Cx!);
    
    internal void UpdateWidth(int widthPixels) => this.AExtents().Cx = UnitConverter.HorizontalPixelToEmu(widthPixels);

    private A.Extents AExtents()
    {
        var aExtents = this.sdkPShapeTreeElement.Descendants<A.Extents>().FirstOrDefault();
        if (aExtents != null)
        {
            return aExtents;
        }

        return new ReferencedPShape(this.sdkTypedOpenXmlPart, this.sdkPShapeTreeElement).ATransform2D().Extents!;
    }
}