using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ShapeCrawler.Assets;
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
    private readonly MemoryStream presStream = new();
    private readonly Stream? inputPresStream;
    private readonly string? inputPresFile;

    /// <summary>
    ///    Opens presentation from the specified stream.
    /// </summary>
    public Presentation(Stream stream)
    {
        this.inputPresStream = stream;
        this.inputPresStream.Position = 0;
        this.inputPresStream.CopyTo(this.presStream);
        
        this.presDocument = PresentationDocument.Open(this.presStream, true);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.SlideMasters = new SlideMasterCollection(this.presDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollection(this.presDocument);
        this.Slides = new UpdatedSlideCollection(
            new SlideCollection(this.presDocument.PresentationPart.SlideParts),
            this.presDocument.PresentationPart);
        this.Footer = new Footer(new UpdatedSlideCollection(
            new SlideCollection(this.presDocument.PresentationPart.SlideParts), this.presDocument.PresentationPart));
        this.Properties =
            this.presDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(this.presDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
    }

    /// <summary>
    ///    Opens presentation from the specified file.
    /// </summary>
    public Presentation(string file)
    {
        this.inputPresFile = file;
        using var fileStream = new FileStream(file, FileMode.Open, FileAccess.Read);
        fileStream.CopyTo(this.presStream);
        
        this.presDocument = PresentationDocument.Open(this.presStream, true);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.SlideMasters = new SlideMasterCollection(this.presDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollection(this.presDocument);
        this.Slides = new UpdatedSlideCollection(
            new SlideCollection(this.presDocument.PresentationPart.SlideParts),
            this.presDocument.PresentationPart);
        this.Footer = new Footer(new UpdatedSlideCollection(
            new SlideCollection(this.presDocument.PresentationPart.SlideParts), this.presDocument.PresentationPart));
        this.Properties =
            this.presDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(this.presDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public Presentation()
    {
        this.presStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("new presentation.pptx");
        
        this.presDocument = PresentationDocument.Open(this.presStream, true);
        this.slideSize = new SlideSize(this.presDocument.PresentationPart!.Presentation.SlideSize!);
        this.SlideMasters = new SlideMasterCollection(this.presDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollection(this.presDocument);
        this.Slides = new UpdatedSlideCollection(
            new SlideCollection(this.presDocument.PresentationPart.SlideParts),
            this.presDocument.PresentationPart);
        this.Footer = new Footer(new UpdatedSlideCollection(
            new SlideCollection(this.presDocument.PresentationPart.SlideParts), this.presDocument.PresentationPart));
        this.Properties =
            this.presDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(this.presDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
        this.Properties.Modified = SCSettings.TimeProvider.UtcNow;
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
    public void Save()
    {
        this.presDocument.Save();
        if (this.inputPresStream is not null)
        {
            this.presDocument.Clone(this.inputPresStream);
        }
        else if (this.inputPresFile is not null)
        {
            var savedPres = this.presDocument.Clone(this.inputPresFile);
            savedPres.Dispose();
        }
    }

    /// <inheritdoc />
    public void Save(Stream stream)
    {
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

    /// <inheritdoc />
    public string AsMarkdown()
    {
        var markdown = new StringBuilder();
        foreach (var slide in this.Slides)
        {
            markdown.AppendLine($"# Slide {slide.Number}");
            var textShapes = slide.Shapes.Where(shape => shape.TextBox is not null && shape.TextBox.Text != string.Empty
                && shape.PlaceholderType != PlaceholderType.SlideNumber);
            var titleShape = textShapes.FirstOrDefault(shape =>
                shape.Name.StartsWith("Title", StringComparison.OrdinalIgnoreCase));
            var nonTitleShapes =
                textShapes.Where(shape => !shape.Name.StartsWith("Title", StringComparison.OrdinalIgnoreCase));
            if (titleShape != null)
            {
                markdown.AppendLine($"## {titleShape.TextBox!.Text}");
            }

            foreach (var shape in nonTitleShapes)
            {
                if (shape.TextBox is not null)
                {
                    markdown.AppendLine(shape.TextBox.Text);
                }
            }

            markdown.AppendLine();
        }

        return markdown.ToString();
    }

    /// <inheritdoc />
    public string AsBase64()
    {
        using var stream = new MemoryStream();
        this.Save(stream);

        return Convert.ToBase64String(stream.ToArray());
    }

    /// <inheritdoc />
    public PresentationDocument GetSDKPresentationDocument() => this.presDocument.Clone();

    /// <summary>
    ///     Releases all resources used by the presentation.
    /// </summary>
    public void Dispose() => this.presDocument.Dispose();

    internal void Validate()
    {
        var nonCriticalErrors = new List<string>
        {
            "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/chart:showDLblsOverMax'.",
            "The element has invalid child element 'http://schemas.microsoft.com/office/drawing/2017/03/chart:dataDisplayOptions16'. " +
            "List of possible elements expected: <http://schemas.microsoft.com/office/drawing/2017/03/chart:dispNaAsBlank>.",
            "The 'uri' attribute is not declared.",
            "The 'mod' attribute is not declared.",
            "The 'mod' attribute is not declared.",
            "The element has unexpected child element 'http://schemas.openxmlformats.org/drawingml/2006/main:noFill'."
        };
        var sdkValidationErrorInfoCollection =
            new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(this.presDocument);
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

        var customErrors = ATableRowErrors(this.presDocument)
            .Concat(ASolidFillErrors(this.presDocument))
            .Concat(sdkErrors);
        if (customErrors.Any())
        {
            var errorMessages = new StringBuilder();
            foreach (var error in customErrors)
            {
                errorMessages.AppendLine(error);
            }

            throw new SCException(errorMessages.ToString());
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
                        throw new SCException("An error occurred while validating the table row structure.");
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