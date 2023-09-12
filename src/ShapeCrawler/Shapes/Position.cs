using System;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed record Position
{
    private readonly Lazy<A.Offset> aOffset;

    internal Position(OpenXmlElement pShapeTreeElement)
    {
        this.aOffset = new Lazy<A.Offset>(() => pShapeTreeElement.Descendants<A.Offset>().First());
    }

    internal int X() => UnitConverter.HorizontalEmuToPixel(this.aOffset.Value.X!);

    internal void UpdateX(int pixels) =>
        this.aOffset.Value.X = new Int64Value(UnitConverter.HorizontalPixelToEmu(pixels));

    internal int Y() => UnitConverter.VerticalEmuToPixel(this.aOffset.Value.Y!);

    internal void UpdateY(int pixels) =>
        this.aOffset.Value.Y = new Int64Value(UnitConverter.VerticalPixelToEmu(pixels));
}