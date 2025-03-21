using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ShapeCrawler.Presentations;
using ShapeCrawler.Slides;
using A = DocumentFormat.OpenXml.Drawing;


#if NETSTANDARD2_0
using ShapeCrawler.Extensions;
#endif

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <inheritdoc />
public sealed class Presentation : IPresentation
{
    private readonly PresentationDocument presDocument;
    private readonly SlideSize slideSize;

    /// <summary>
    ///    Opens presentation from the specified stream.
    /// </summary>
    public Presentation(Stream stream)
    {
        this.presDocument = PresentationDocument.Open(stream, true);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.SlideMasters = new SlideMasterCollection(this.presDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollectionCollection(this.presDocument);
        this.Slides = new SlideCollection(this.presDocument.PresentationPart);
        this.Footer = new Footer(new SlideCollection(this.presDocument.PresentationPart));
        this.Properties =
            new PresentationProperties(this.presDocument.CoreFilePropertiesPart!.OpenXmlPackage.PackageProperties);
    }

    /// <summary>
    ///    Opens presentation from the specified file.
    /// </summary>
    public Presentation(string file)
    {
        this.presDocument = PresentationDocument.Open(file, true);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.SlideMasters = new SlideMasterCollection(this.presDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollectionCollection(this.presDocument);
        this.Slides = new SlideCollection(this.presDocument.PresentationPart);
        this.Footer = new Footer(new SlideCollection(this.presDocument.PresentationPart));
        this.Properties =
            new PresentationProperties(this.presDocument.CoreFilePropertiesPart!.OpenXmlPackage.PackageProperties);
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public Presentation()
    {
        var assets = new Assets(Assembly.GetExecutingAssembly());
        var stream = assets.StreamOf("new-presentation.pptx");

        this.presDocument = PresentationDocument.Open(stream, true);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.SlideMasters = new SlideMasterCollection(this.presDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollectionCollection(this.presDocument);
        this.Slides = new SlideCollection(this.presDocument.PresentationPart);
        this.Footer = new Footer(new SlideCollection(this.presDocument.PresentationPart));
        this.Properties =
            new PresentationProperties(this.presDocument.CoreFilePropertiesPart!.OpenXmlPackage.PackageProperties)
            {
                Modified = SCSettings.TimeProvider.UtcNow
            };
    }

    internal Presentation(PresentationDocument presDocument)
    {
        this.presDocument = presDocument;
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.SlideMasters = new SlideMasterCollection(this.presDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollectionCollection(this.presDocument);
        this.Slides = new SlideCollection(this.presDocument.PresentationPart);
        this.Footer = new Footer(new SlideCollection(this.presDocument.PresentationPart));
        this.Properties =
            new PresentationProperties(this.presDocument.CoreFilePropertiesPart!.OpenXmlPackage.PackageProperties);
    }

    /// <inheritdoc />
    public ISlideCollection Slides { get; }

    /// <inheritdoc />
    public decimal SlideHeight
    {
        get => this.slideSize.Height;
        set => this.slideSize.Height = value;
    }

    /// <inheritdoc />
    public decimal SlideWidth
    {
        get => this.slideSize.Width;
        set => this.slideSize.Width = value;
    }

    /// <inheritdoc />
    public ISlideMasterCollection SlideMasters { get; }

    /// <inheritdoc />
    public ISectionCollection Sections { get; }

    /// <inheritdoc />
    public IFooter Footer { get; }

    /// <inheritdoc />
    public IPresentationProperties Properties { get; }

    /// <inheritdoc />
    public ISlide Slide(int number) => this.Slides[number - 1];

    /// <summary>
    ///     Gets Slide Master by number.
    /// </summary>
    public ISlideMaster SlideMaster(int number) => this.SlideMasters[number - 1];

    /// <inheritdoc />
    public void Save() => this.presDocument.Save();

    /// <inheritdoc />
    public void Save(Stream stream)
    {
        this.Save();
        this.Properties.Modified = SCSettings.TimeProvider.UtcNow;

        if (stream is FileStream fileStream)
        {
            var mStream = new MemoryStream();
            this.presDocument.Clone(mStream);
            mStream.Position = 0;
            mStream.CopyTo(fileStream);
        }
        else
        {
            this.presDocument.Clone(stream);
        }
    }

    /// <inheritdoc />
    public void Save(string file)
    {
        this.Save();
        using var stream = new FileStream(file, FileMode.Create);
        this.Save(stream);
    }
    
    /// <summary>
    ///     Releases all resources used by the presentation.
    /// </summary>
    public void Dispose() => this.presDocument.Dispose();

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

    private static IEnumerable<string> ValidateATableRows(PresentationDocument presDocument)
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
                if (element is A.TableCell)
                {
                    lastTableCellIndex = i;
                }
                else if (element is A.ExtensionList)
                {
                    extListIndex = i;
                }
            }

            if (extListIndex < lastTableCellIndex)
            {
                yield return
                    "Invalid table row structure: ExtensionList element must appear after all TableCell elements in a TableRow";
            }
        }
    }

    private static IEnumerable<string> ValidateASolidFill(PresentationDocument presDocument)
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