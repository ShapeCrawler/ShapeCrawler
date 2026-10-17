using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class ShapePropertiesExtensions
{
    extension(OpenXmlCompositeElement pShapeProperties)
    {
        internal void AddNoFill()
        {
            pShapeProperties.RemoveAllChildren<A.GradientFill>();
            pShapeProperties.RemoveAllChildren<A.SolidFill>();
            pShapeProperties.RemoveAllChildren<A.PatternFill>();
            pShapeProperties.RemoveAllChildren<A.NoFill>();
            pShapeProperties.RemoveAllChildren<A.BlipFill>();
            pShapeProperties.RemoveAllChildren<A.GroupFill>();

            pShapeProperties.AddChild(new A.NoFill(), true);
        }

        internal void AddSolidFill(string hex)
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

            var aRgbColorModelHex = new A.RgbColorModelHex { Val = hex };

            aSolidFill.Append(aRgbColorModelHex);
        }
    }
}
