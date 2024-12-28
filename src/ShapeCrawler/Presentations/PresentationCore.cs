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
        this.FileProperties.Modified = ShapeCrawlerInternal.TimeProvider.UtcNow;
        var cloned = this.sdkPresDocument.Clone(path);
        cloned.Dispose();
    }

    internal void CopyTo(Stream stream)
    {
        this.FileProperties.Modified = ShapeCrawlerInternal.TimeProvider.UtcNow;
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
        var nonCriticalErrorDesc = new List<string>
        {
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
                "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
                "The 'uri' attribute is not declared.",
                "The 'mod' attribute is not declared.",
                "The 'mod' attribute is not declared.",
                "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:noFill'."
        };
        var sdkErrors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(this.sdkPresDocument);
        sdkErrors = sdkErrors.Where(errorInfo => !nonCriticalErrorDesc.Contains(errorInfo.Description));
        sdkErrors = sdkErrors.DistinctBy(x => new { x.Description, x.Path?.XPath }).ToList();

        if (sdkErrors.Any())
        {
            throw new SCException("Presentation is invalid.");
        }
        
        var errors = this.ValidateATableRows(this.sdkPresDocument);
        errors = errors.Concat(this.ValidateASolidFill(this.sdkPresDocument));
        if (errors.Any())
        {
            throw new SCException("Presentation is invalid.");
        }
    }

    private IEnumerable<string> ValidateATableRows(PresentationDocument presDocument)
    {
        var aTableRows = presDocument.PresentationPart!.SlideParts
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
    
    private IEnumerable<string> ValidateASolidFill(PresentationDocument presDocument)
    {
        var aText = presDocument.PresentationPart!.SlideParts
            .SelectMany(slidePart => slidePart.Slide.Descendants<A.Text>());
        aText = aText.Concat(presDocument.PresentationPart!.SlideMasterParts
            .SelectMany(slidePart => slidePart.SlideMaster.Descendants<A.Text>())).ToList();

        foreach (var text in aText)
        {
            var runProperties = text.Parent!.GetFirstChild<A.RunProperties>();
            
            if ((runProperties?.Descendants<A.SolidFill>()?.Any() ?? false) 
                && runProperties.ChildElements.Take(2).All(x => x is not A.SolidFill))
            {
                yield return $"Invalid solid fill structure: SolidFill element must be index 0";
            }
        }
    }
}