using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace ShapeCrawler.Tests.Helpers;

public static class PresentationDocumentExtensions
{
    public static bool IsValid(this PresentationDocument pres)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(pres);

        return !errors.Any();
    }
}