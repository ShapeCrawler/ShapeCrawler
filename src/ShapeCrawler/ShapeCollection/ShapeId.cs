using System.Linq;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.ShapeCollection;

internal class ShapeId
{
    private readonly OpenXmlElement pShapeTreeElement;

    internal ShapeId(OpenXmlElement pShapeTreeElement)
    {
        this.pShapeTreeElement = pShapeTreeElement;
    }

    internal int Value()
    {
        var pCNvPr = this.pShapeTreeElement.Descendants<P.NonVisualDrawingProperties>().First();
        return (int)pCNvPr.Id!.Value!;
    }

    internal void Update(int id)
    {
        var pCNvPr = this.pShapeTreeElement.Descendants<P.NonVisualDrawingProperties>().First();
        pCNvPr.Id!.Value = (uint)id;
    }
}