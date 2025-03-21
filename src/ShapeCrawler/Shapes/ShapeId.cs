using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal class ShapeId(OpenXmlElement pShapeTreeElement)
{
    internal int Value()
    {
        var pCNvPr = pShapeTreeElement.Descendants<P.NonVisualDrawingProperties>().First();
        return (int)pCNvPr.Id!.Value;
    }

    internal void Update(int id)
    {
        var pCNvPr = pShapeTreeElement.Descendants<P.NonVisualDrawingProperties>().First();
        pCNvPr.Id!.Value = (uint)id;
    }
}