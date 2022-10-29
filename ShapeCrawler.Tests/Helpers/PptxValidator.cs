using System.Collections.Generic;
using System.Diagnostics;
using SharpCompress.Readers;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using SharpCompress.Common;

namespace ShapeCrawler.Tests.Helpers;

public static class PptxValidator
{
    private static List<ValidationError> warnings = new List<ValidationError>
    {
        new(
            "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
            "/c:chartSpace[1]/c:chart[1]"),
        new("/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]", "/c:chartSpace[1]/c:chart[1]"),
        new(
            "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
            "/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]"),
        new(
            "The 'uri' attribute is not declared.",
            "/c:chartSpace[1]/c:chart[1]/c:extLst[1]/c:ext[1]"),
    };

    public static List<ValidationErrorInfo> Validate(IPresentation pres)
    {
        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
        var errors = validator.Validate(pres.SDKPresentation);

        var removing = new List<ValidationErrorInfo>();
        foreach (var error in errors)
        {
            if (warnings.Any(x => x.Description == error.Description && x.Path == error.Path?.XPath))
            {
                removing.Add(error);
            }
        }

        errors = errors.Except(removing).DistinctByCustom(x=> new {x.Description, x.Path?.XPath}).ToList();
        
        return errors.ToList();
    }
}

public class ValidateResponse
{
    public bool IsValid { get; }
    public string ErrorMessage { get; }

    public ValidateResponse()
    {
        IsValid = true;
    }

    public ValidateResponse(string errorMessage)
    {
        ErrorMessage = errorMessage;
    }
}