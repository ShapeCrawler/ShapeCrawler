using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class ShapePropertiesExtensions
{
    internal static A.SolidFill AddASolidFill(this P.ShapeProperties pShapeProperties, string hex)
    {
        var aSolidFill = pShapeProperties.GetFirstChild<A.SolidFill>();
        aSolidFill?.Remove();
        var aRgbColorModelHex = new A.RgbColorModelHex
        {
            Val = hex
        };
        aSolidFill = new A.SolidFill();
        aSolidFill.Append(aRgbColorModelHex);
        pShapeProperties.Append(aSolidFill);

        return aSolidFill;
    }

    internal static void AddAXfrm(this P.ShapeProperties pSpPr, long xEmu, long yEmu, long wEmu, long hEmu)
    {
        var aXfrm = pSpPr.Transform2D;
        aXfrm?.Remove();
        
        aXfrm = new A.Transform2D();
        pSpPr.Append(aXfrm);

        var aOff = new A.Offset
        {
            X = xEmu,
            Y = yEmu
        };
        aXfrm.Append(aOff);
        
        var aExt = new A.Extents
        {
            Cx = wEmu,
            Cy = hEmu
        };
        aXfrm.Append(aExt);
    }
}