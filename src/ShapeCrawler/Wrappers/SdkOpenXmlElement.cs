using System.Linq;
using DocumentFormat.OpenXml;

namespace ShapeCrawler.Wrappers;

/// <summary>
///     Wrapper for <see cref="DocumentFormat.OpenXml.OpenXmlElement"/>.
/// </summary>
internal sealed record SdkOpenXmlElement
{
    private readonly OpenXmlElement openXmlElement;

    internal SdkOpenXmlElement(OpenXmlElement openXmlElement)
    {
        this.openXmlElement = openXmlElement;
    }

    internal T FirstAncestor<T>() where T : OpenXmlElement
    {
        return this.openXmlElement.Ancestors<T>().First();
    }
}