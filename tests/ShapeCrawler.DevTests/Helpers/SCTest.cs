using System.Reflection;
using System.Text;
using System.Xml.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Fixture;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.DevTests.Helpers;

public abstract class SCTest
{
    protected readonly Fixtures fixture = new();

    protected static T GetWorksheetCellValue<T>(byte[] workbookByteArray, string cellAddress)
    {
        var stream = new MemoryStream(workbookByteArray);
        var xlWorkbook = new XLWorkbook(stream);
        var cellValue = xlWorkbook.Worksheets.First().Cell(cellAddress).Value;

        return (T)cellValue;
    }

    public static MemoryStream TestAsset(string file)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var stream = assembly.GetResourceStream(file);
        var mStream = new MemoryStream();
        stream.CopyTo(mStream);
        mStream.Position = 0;

        return mStream;
    }

    protected static string StringOf(string fileName)
    {
        var stream = TestAsset(fileName);
        return System.Text.Encoding.UTF8.GetString(stream.ToArray());
    }

    protected static Presentation SaveAndOpenPresentation(IPresentation presentation)
    {
        var stream = new MemoryStream();
        presentation.Save(stream);

        return new Presentation(stream);
    }

    protected static PresentationDocument SaveAndOpenPresentationAsSdk(IPresentation presentation)
    {
        var stream = new MemoryStream();
        presentation.Save(stream);
        stream.Position = 0;

        return PresentationDocument.Open(stream, true);
    }

    protected static void ValidatePresentation(Presentation pres)
    {
        var presDocument = pres.GetSDKPresentationDocument();
        var nonCriticalErrors = new List<string>
        {
            "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
            "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. " +
            "List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
            "The 'uri' attribute is not declared.",
            "The 'mod' attribute is not declared.",
            "The 'mod' attribute is not declared.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:noFill'.",
            "The element has unexpected child element 'http://schemas.microsoft.com/office/drawing/2012/chart:leaderLines'."
        };
        var sdkValidationErrorInfoCollection =
            new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(presDocument);
        sdkValidationErrorInfoCollection =
            sdkValidationErrorInfoCollection.Where(errorInfo => !nonCriticalErrors.Contains(errorInfo.Description));
        sdkValidationErrorInfoCollection =
        [
            .. sdkValidationErrorInfoCollection.DistinctBy(errorInfo =>
                new { errorInfo.Description, errorInfo.Path?.XPath })
        ];
        var sdkErrors = new List<string>();
        foreach (var validationErrorInfo in sdkValidationErrorInfoCollection)
        {
            var xmlError = new XElement("error");
            xmlError.Add(new XElement("id", validationErrorInfo.Id));
            xmlError.Add(new XElement("description", validationErrorInfo.Description));
            xmlError.Add(new XElement("xpath", validationErrorInfo.Path?.XPath));
            sdkErrors.Add(xmlError.ToString());
        }

        var customErrors = ATableRowErrors(presDocument)
            .Concat(ASolidFillErrors(presDocument))
            .Concat(sdkErrors);
        if (customErrors.Any())
        {
            var errorMessages = new StringBuilder();
            foreach (var error in customErrors)
            {
                errorMessages.AppendLine(error);
            }

            throw new Exception(errorMessages.ToString());
        }
    }

    private static IEnumerable<string> ATableRowErrors(PresentationDocument presDocument)
    {
        var aTableRows = presDocument.PresentationPart!.SlideParts
            .SelectMany(slidePart => slidePart.Slide.Descendants<A.TableRow>());

        foreach (var aTableRow in aTableRows)
        {
            var aExtLst = aTableRow.GetFirstChild<A.ExtensionList>();
            if (aExtLst == null)
            {
                continue;
            }

            var lastTableCellIndex = -1;
            var extListIndex = -1;

            for (int i = 0; i < aTableRow.ChildElements.Count; i++)
            {
                var element = aTableRow.ChildElements[i];
                switch (element)
                {
                    case A.TableCell:
                        lastTableCellIndex = i;
                        break;
                    case A.ExtensionList:
                        extListIndex = i;
                        break;
                    default:
                        throw new Exception("An error occurred while validating the table row structure.");
                }
            }

            if (extListIndex < lastTableCellIndex)
            {
                yield return
                    "Invalid table row structure: ExtensionList element must appear after all TableCell elements in a TableRow";
            }
        }
    }

    private static IEnumerable<string> ASolidFillErrors(PresentationDocument presDocument)
    {
        var aText = presDocument.PresentationPart!.SlideParts
            .SelectMany(slidePart => slidePart.Slide.Descendants<A.Text>());
        aText =
        [
            .. aText,
            .. presDocument.PresentationPart!.SlideMasterParts
                .SelectMany(slidePart => slidePart.SlideMaster.Descendants<A.Text>()),
        ];

        foreach (var text in aText)
        {
            var runProperties = text.Parent!.GetFirstChild<A.RunProperties>();
            if ((runProperties?.Descendants<A.SolidFill>().Any() ?? false)
                && runProperties.ChildElements.Take(2).All(x => x is not A.SolidFill))
            {
                yield return "Invalid solid fill structure: SolidFill element must be index 0";
            }
        }
    }
}