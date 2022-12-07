using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;

namespace ShapeCrawler.Tests.Helpers;

public static class PptxValidator
{
    private static readonly List<ValidationError> NonCritical = new()
    {
        new ValidationError(
            "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
            "/c:chartSpace[1]/c:chart[1]"),
        new ValidationError("/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]", "/c:chartSpace[1]/c:chart[1]"),
        new ValidationError(
            "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
            "/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]"),
        new ValidationError(
            "The 'uri' attribute is not declared.",
            "/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]"),
        new ValidationError(
        
            "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:pPr'.",
            "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]/a:p[1]")
    };

    public static List<ValidationErrorInfo> Validate(IPresentation pres)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(pres.SDKPresentation);

        var removing = new List<ValidationErrorInfo>();
        foreach (var error in errors)
        {
            if (NonCritical.Any(x => x.Description == error.Description && x.Path == error.Path?.XPath))
            {
                removing.Add(error);
            }
        }

        errors = errors.Except(removing).DistinctByCustom(x=> new {x.Description, x.Path?.XPath}).ToList();
        
        return errors.ToList();
    }
}