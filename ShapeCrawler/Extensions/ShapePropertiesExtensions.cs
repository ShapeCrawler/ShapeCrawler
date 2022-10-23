using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class ShapePropertiesExtensions
{
    internal static A.SolidFill AddASolidFill(this P.ShapeProperties pShapeProperties, string hex)
    {
        var aSolidFill = pShapeProperties.GetFirstChild<A.SolidFill>();
        var aNoFill = pShapeProperties.GetFirstChild<A.NoFill>();
        aSolidFill?.Remove();
        aNoFill?.Remove();

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