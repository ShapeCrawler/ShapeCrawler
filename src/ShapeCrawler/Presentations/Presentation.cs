using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

#if NETSTANDARD2_0
using ShapeCrawler.Extensions;
#endif

namespace ShapeCrawler;

/// <inheritdoc />
public sealed class Presentation : IPresentation
{
    private readonly PresentationDocument presDocument;
    private readonly SlideSize slideSize;

    public Presentation(Stream stream)
    {
        this.presDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.presDocument.PresentationPart!.SlideMasterParts;
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.presDocument);
        this.Slides = new SlideCollection(this.presDocument.PresentationPart);
        this.Footer = new Footer(this);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.Metadata = new FileProperties(this.presDocument.CoreFilePropertiesPart!);
    }
    
    public Presentation()
    {
        var assets = new Assets(Assembly.GetExecutingAssembly());
        var stream = assets.StreamOf("new-presentation.pptx");
        
        this.presDocument = PresentationDocument.Open(stream, true);
        var sdkMasterParts = this.presDocument.PresentationPart!.SlideMasterParts;
        this.SlideMasters = new SlideMasterCollection(sdkMasterParts);
        this.Sections = new Sections(this.presDocument);
        this.Slides = new SlideCollection(this.presDocument.PresentationPart);
        this.Footer = new Footer(this);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.Metadata = new FileProperties(this.presDocument.CoreFilePropertiesPart!);
    }

    /// <inheritdoc />
    public ISlideCollection Slides { get; }

    /// <inheritdoc />
    public decimal SlideHeight
    {
        get => this.slideSize.Height();
        set => this.slideSize.UpdateHeight(value);
    }

    /// <inheritdoc />
    public decimal SlideWidth
    {
        get => this.slideSize.Width();
        set => this.slideSize.UpdateWidth(value);
    }

    /// <inheritdoc />
    public ISlideMasterCollection SlideMasters { get; }

    /// <inheritdoc />
    public ISections Sections { get; }

    /// <inheritdoc />
    public IFooter Footer { get; }
    
    /// <inheritdoc />
    public IPresentationMetadata Metadata { get; }
    
    /// <inheritdoc />
    public ISlide Slide(int number) => this.Slides[number - 1];

    /// <inheritdoc />
    public void Save() => this.presDocument.Save();
    
    /// <inheritdoc />
    public void Copy(Stream stream)
    {
        this.Metadata.Modified = SCSettings.TimeProvider.UtcNow;
        this.presDocument.Clone(stream);
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
        var sdkErrors = new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(this.presDocument);
        sdkErrors = sdkErrors.Where(errorInfo => !nonCriticalErrors.Contains(errorInfo.Description));
        sdkErrors = [.. sdkErrors.DistinctBy(errorInfo => new { errorInfo.Description, errorInfo.Path?.XPath })];

        if (sdkErrors.Any())
        {
            throw new SCException("Presentation is invalid.");
        }
        
        var errors = ValidateATableRows(this.presDocument);
        errors = errors.Concat(ValidateASolidFill(this.presDocument));
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
        aText =
        [
            .. aText,
            .. sdkPres.PresentationPart!.SlideMasterParts
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