using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class ShapePropertiesExtensions
{
    internal static void AddASolidFill(this OpenXmlCompositeElement pShapeProperties, string hex)
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
            var aOutline = pShapeProperties.GetFirstChild<A.Outline>();
            if (aOutline != null)
            {
                pShapeProperties.InsertBefore(aSolidFill, aOutline);
            }
            else
            {
                pShapeProperties.Append(aSolidFill);
            }
        }
        
        var aRgbColorModelHex = new A.RgbColorModelHex
        {
            Val = hex
        };
        
        aSolidFill.Append(aRgbColorModelHex);
    }

    internal static A.Outline AddAOutline(this OpenXmlCompositeElement pSpPr)
    {
        var aOutline = pSpPr.GetFirstChild<A.Outline>();
        aOutline?.Remove();
        
        var aSchemeClr = new A.SchemeColor { Val = new EnumValue<A.SchemeColorValues>(A.SchemeColorValues.Text1) };
        var aSolidFill = new A.SolidFill(aSchemeClr);
        var aOutlineNew = new A.Outline(aSolidFill);
        pSpPr.Append(aOutlineNew);

        return aOutlineNew;
    }
}