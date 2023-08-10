using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed record ShapeLocation
{
    private readonly A.Offset aOffset;

    internal ShapeLocation(A.Offset aOffset)
    {
        this.aOffset = aOffset;
    }

    internal int X()
    {
        return UnitConverter.HorizontalEmuToPixel(this.aOffset.X!);
    }

    internal void UpdateX(int pixels)
    {
        this.aOffset.X = new Int64Value(UnitConverter.HorizontalPixelToEmu(pixels));
    }

    internal int Y()
    {
        return UnitConverter.VerticalEmuToPixel(this.aOffset.Y!);
    }

    internal void UpdateY(int pixels)
    {
        this.aOffset!.Y = new Int64Value(UnitConverter.VerticalPixelToEmu(pixels));
    }
}