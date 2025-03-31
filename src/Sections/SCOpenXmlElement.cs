using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Sections;

internal readonly ref struct SCOpenXmlElement(OpenXmlElement openXmlElement)
{
    internal PresentationDocument PresentationDocument =>
        (PresentationDocument)openXmlElement.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!
            .OpenXmlPackage;
}