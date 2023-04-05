using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class ShapePropertiesExtensions
{
    internal static A.SolidFill AddASolidFill(this TypedOpenXmlCompositeElement pShapeProperties, string hex)
    {
        pShapeProperties.GetFirstChild<A.GradientFill>()?.Remove();
        pShapeProperties.GetFirstChild<A.PatternFill>()?.Remove();
        pShapeProperties.GetFirstChild<A.NoFill>()?.Remove();
        pShapeProperties.GetFirstChild<A.BlipFill>()?.Remove();
        
        var aSolidFill = pShapeProperties.GetFirstChild<A.SolidFill>();
        if (aSolidFill != null)
        {
            foreach (var child in aSolidFill)
            {
                child.Remove();
            }
        }
        else
        {
            aSolidFill = new A.SolidFill();
            pShapeProperties.Append(aSolidFill);
        }
        
        var aRgbColorModelHex = new A.RgbColorModelHex
        {
            Val = hex
        };
        
        aSolidFill.Append(aRgbColorModelHex);

        return aSolidFill;
    }

    internal static A.Outline AddAOutline(this P.ShapeProperties pSpPr)
    {
        var aOutline = pSpPr.GetFirstChild<A.Outline>();
        aOutline?.Remove();
        
        var aSchemeClr = new A.SchemeColor { Val = new EnumValue<A.SchemeColorValues>(A.SchemeColorValues.Text1) };
        var aSolidFill = new A.SolidFill(aSchemeClr);
        var aOutlineNew = new A.Outline(aSolidFill);
        pSpPr.Append(aOutlineNew);

        return aOutlineNew;
    }
    
    internal static A.Transform2D AddAXfrm(this P.ShapeProperties pSpPr, long xEmu, long yEmu, long wEmu, long hEmu)
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

        return aXfrm;
    }
}