using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

#if NETSTANDARD2_0
using ShapeCrawler.Extensions;
#endif

namespace ShapeCrawler.Presentations;

internal sealed class PresentationCore
{
    private readonly PresentationDocument sdkPresDocument;
    private readonly SlideSize slideSize;

    internal PresentationCore(byte[] bytes)
    {
        var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        this.sdkPresDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.sdkPresDocument.PresentationPart!.SlideMasterParts;    
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.sdkPresDocument);
        this.Slides = new Slides(this.sdkPresDocument.PresentationPart);
        this.Footer = new Footer(this);
        this.slideSize = new SlideSize(this.sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
        this.FileProperties = new(this.sdkPresDocument.CoreFilePropertiesPart!);
    }

    internal PresentationCore(Stream stream)
    {
        stream.Position = 0;
        this.sdkPresDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.sdkPresDocument.PresentationPart!.SlideMasterParts;
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.sdkPresDocument);
        this.Slides = new Slides(this.sdkPresDocument.PresentationPart);
        this.Footer = new Footer(this);
        this.slideSize = new SlideSize(this.sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
        this.FileProperties = new(this.sdkPresDocument.CoreFilePropertiesPart!);
    }

    internal ISlides Slides { get; }

    internal decimal SlideHeight
    {
        get => this.slideSize.Height();
        set => this.slideSize.UpdateHeight(value);
    }

    internal decimal SlideWidth
    {
        get => this.slideSize.Width();
        set => this.slideSize.UpdateWidth(value);
    }

    internal ISlideMasterCollection SlideMasters { get; }

    internal ISections Sections { get; }

    internal IFooter Footer { get; }

    internal FileProperties FileProperties { get; }

    internal void CopyTo(string path)
    {
        this.FileProperties.Modified = SCSettings.TimeProvider.UtcNow;
        var cloned = this.sdkPresDocument.Clone(path);
        cloned.Dispose();
    }

    internal void CopyTo(Stream stream)
    {
        this.FileProperties.Modified = SCSettings.TimeProvider.UtcNow;
        this.sdkPresDocument.Clone(stream);
    }

    internal byte[] AsByteArray()
    {
        var stream = new MemoryStream();
        this.sdkPresDocument.Clone(stream);

        return stream.ToArray();
    }

    internal void Validate()
    {
        var nonCriticalErrors = new List<string>
        {
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
                "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
                "The 'uri' attribute is not declared.",
                "The 'mod' attribute is not declared.",
                "The 'mod' attribute is not declared.",
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:noFill'."
        };
        var sdkErrors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(this.sdkPresDocument);
        sdkErrors = sdkErrors.Where(errorInfo => !nonCriticalErrors.Contains(errorInfo.Description));
        sdkErrors = sdkErrors.DistinctBy(errorInfo => new { errorInfo.Description, errorInfo.Path?.XPath }).ToList();

        if (sdkErrors.Any())
        {
            throw new SCException("Presentation is invalid.");
        }
        
        var errors = ValidateATableRows(this.sdkPresDocument);
        errors = errors.Concat(ValidateASolidFill(this.sdkPresDocument));
        if (errors.Any())
        {
            throw new SCException("Presentation is invalid.");
        }
    }

    private static IEnumerable<string> ValidateATableRows(PresentationDocument sdkPres)
    {
        var aTableRows = sdkPres.PresentationPart!.SlideParts
            .SelectMany(slidePart => slidePart.Slide.Descendants<A.TableRow>());

        foreach (var aTableRow in aTableRows)
        {
            var aExtLst = aTableRow.GetFirstChild<A.ExtensionList>();
            if (aExtLst != null)
            {
                var lastTableCellIndex = -1;
                var extListIndex = -1;

                // Find indices of last TableCell and ExtensionList
                for (int i = 0; i < aTableRow.ChildElements.Count; i++)
                {
                    var element = aTableRow.ChildElements[i];
                    if (element is A.TableCell)
                    {
                        lastTableCellIndex = i;
                    }
                    else if (element is A.ExtensionList)
                    {
                        extListIndex = i;
                    }
                }

                // If ExtensionList appears before the last TableCell, yield the error
                if (extListIndex < lastTableCellIndex)
                {
                    yield return "Invalid table row structure: ExtensionList element must appear after all TableCell elements in a TableRow";
                }
            }
        }
    }
    
    private static IEnumerable<string> ValidateASolidFill(PresentationDocument sdkPres)
    {
        var aText = sdkPres.PresentationPart!.SlideParts
            .SelectMany(slidePart => slidePart.Slide.Descendants<A.Text>());
        aText = aText.Concat(sdkPres.PresentationPart!.SlideMasterParts
            .SelectMany(slidePart => slidePart.SlideMaster.Descendants<A.Text>())).ToList();

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