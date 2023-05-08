using DocumentFormat.OpenXml.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class RunPropertiesExtensions
{
    internal static void AddAHighlight(this RunProperties arPr, string? hex)
    {
        var aHighlight = arPr.GetFirstChild<A.Highlight>();
        aHighlight?.Remove();

        // Don't add a new node
        if (hex is null)
        {
            return;
        }
        
        var aSrgbClr = new A.RgbColorModelHex
        {
            Val = hex
        };
        aHighlight = new A.Highlight();
        aHighlight.Append(aSrgbClr);
        
        arPr.Append(aHighlight);
    } 
}