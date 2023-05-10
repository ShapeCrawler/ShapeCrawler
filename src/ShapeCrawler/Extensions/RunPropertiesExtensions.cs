using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class RunPropertiesExtensions
{
    internal static void AddAHighlight(this RunProperties arPr, string hex)
    {
        var aHighlight = arPr.GetFirstChild<A.Highlight>();
        aHighlight?.Remove();

        var aSrgbClr = new A.RgbColorModelHex
        {
            Val = hex
        };
        aHighlight = new A.Highlight();
        aHighlight.Append(aSrgbClr);

        arPr.Append(aHighlight);
    }

    internal static void AddAHighlight(this RunProperties arPr, SCColor hex)
    {
        var aHighlight = arPr.GetFirstChild<A.Highlight>();
        aHighlight?.Remove();

        if (hex.IsTransparent)
        {
            return;
        }

        aHighlight = new A.Highlight();
        aHighlight.Append(hex.ToRgbColorModelHex());

        arPr.Append(aHighlight);
    }

    internal static A.RgbColorModelHex ToRgbColorModelHex(this SCColor color)
    {
        // Initialize color model.
        var model = new A.RgbColorModelHex
        {
            Val = color.ToString(),
        };

        // Solid color doesn't have alpha value.
        if (color.IsSolid)
        {
            // Solid colores doesn't need to specify alpha value.
            return model;
        }

        // Creates a alpha node...
        var alpha = new A.Alpha
        {
            Val = (Int32Value)(100000f * (color.Alpha / SCColor.OPACITY))
        };

        model.AddChild(alpha);

        return model;
    }
}