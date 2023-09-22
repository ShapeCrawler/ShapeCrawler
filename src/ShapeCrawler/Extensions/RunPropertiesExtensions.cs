using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class RunPropertiesExtensions
{
    internal static void AddAHighlight(this RunProperties arPr, Color color)
    {
        var aHighlight = arPr.GetFirstChild<A.Highlight>();
        aHighlight?.Remove();

        if (color.IsTransparent)
        {
            return;
        }

        aHighlight = new A.Highlight();
        aHighlight.Append(color.ToRgbColorModelHex());

        arPr.Append(aHighlight);
    }

    private static A.RgbColorModelHex ToRgbColorModelHex(this Color color)
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
            Val = (Int32Value)(100000f * (color.Alpha / Color.OPACITY))
        };

        model.AddChild(alpha);

        return model;
    }
}