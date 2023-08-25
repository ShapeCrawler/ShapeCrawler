using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed record ShapeSize
{
    private readonly A.Extents aExtents;

    internal ShapeSize(A.Extents aExtents)
    {
        this.aExtents = aExtents;
    }

    internal int Height() => UnitConverter.VerticalEmuToPixel(this.aExtents.Cy!);

    internal void UpdateHeight(int heightPixels) => this.aExtents.Cx = UnitConverter.VerticalPixelToEmu(heightPixels);
    internal int Width() => UnitConverter.HorizontalEmuToPixel(this.aExtents.Cx!);

    internal void UpdateWidth(int widthPixels) => this.aExtents.Cx = UnitConverter.HorizontalPixelToEmu(widthPixels);
}