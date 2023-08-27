using System.Linq;
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Wrappers;

internal sealed record SDKATypedOpenXmlCompositeElementWrap
{
    private readonly TypedOpenXmlCompositeElement typedOpenXmlCompositeElement;

    internal SDKATypedOpenXmlCompositeElementWrap(TypedOpenXmlCompositeElement typedOpenXmlCompositeElement)
    {
        this.typedOpenXmlCompositeElement = typedOpenXmlCompositeElement;
    }

    internal T FirstAncestor<T>() where T : OpenXmlElement
    {
        return this.typedOpenXmlCompositeElement.Ancestors<T>().First();
    }
}