using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Sections;

internal readonly ref struct SCOpenXmlLeafElement(SectionSlideIdListEntry sectionSlideIdListEntry)
{
    internal PresentationDocument PresentationDocument =>
        (PresentationDocument)(sectionSlideIdListEntry.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!
            .OpenXmlPackage);
}