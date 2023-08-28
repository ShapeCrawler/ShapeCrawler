using System.Linq;
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Wrappers;

internal sealed record SDKOpenXmlElementWrap
{
    private readonly OpenXmlElement openXmlElement;

    internal SDKOpenXmlElementWrap(OpenXmlElement openXmlElement)
    {
        this.openXmlElement = openXmlElement;
    }

    internal T FirstAncestor<T>() where T : OpenXmlElement
    {
        return this.openXmlElement.Ancestors<T>().First();
    }
}