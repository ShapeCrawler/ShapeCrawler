using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class ShapePropertiesExtensions
{
    internal static A.SolidFill AddASolidFill(this P.ShapeProperties pShapeProperties, string hex)
    {
        var aSolidFill = pShapeProperties.GetFirstChild<A.SolidFill>();
        if (aSolidFill is not null)
        {
            aSolidFill.Remove();
        }

        var aRgbColorModelHex = new A.RgbColorModelHex
        {
            Val = hex
        };
        aSolidFill = new A.SolidFill();
        aSolidFill.Append(aRgbColorModelHex);
        pShapeProperties.Append(aSolidFill);

        return aSolidFill;
    }
}