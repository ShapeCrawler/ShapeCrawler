using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;

namespace ShapeCrawler.Shapes;

internal sealed class SCShapeLocation
{
    private readonly DocumentFormat.OpenXml.Drawing.Transform2D aTransform2D;

    internal SCShapeLocation(DocumentFormat.OpenXml.Drawing.Transform2D aTransform2D)
    {
        this.aTransform2D = aTransform2D;
    }

    internal int ParseX()
    {
        return UnitConverter.HorizontalEmuToPixel(this.aTransform2D.Offset!.X!);
    }

    internal void UpdateX(int pixels)
    {
        this.aTransform2D.Offset!.X = new Int64Value(UnitConverter.HorizontalPixelToEmu(pixels));
    }

    internal int ParseY()
    {
        return UnitConverter.VerticalEmuToPixel(this.aTransform2D.Offset!.Y!);
    }

    internal void UpdateY(int pixels)
    {
        this.aTransform2D.Offset!.Y = new Int64Value(UnitConverter.VerticalPixelToEmu(pixels));
    }
}