using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler;

internal readonly ref struct SCOpenXmlElement(OpenXmlElement openXmlElement)
{
    internal PresentationDocument ParentPresentationDocument => (PresentationDocument)this.ParentOpenXmlPart.OpenXmlPackage;

    internal OpenXmlPart ParentOpenXmlPart => openXmlElement.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
}