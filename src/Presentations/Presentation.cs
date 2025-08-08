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
using P = DocumentFormat.OpenXml.Presentation;
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

    /// <summary>
    ///     Creates a new presentation using fluent configuration.
    /// </summary>
    public Presentation(Action<DraftPresentation> configure)
        : this()
    {
        var draft = new DraftPresentation();
        configure(draft);
        draft.ApplyTo(this);
    }

    // Note: Create(Action<DraftPresentation>) is defined once earlier in the class.

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
    
    /// <summary>
    ///     Starts a fluent creation of a new presentation.
    /// </summary>
    public static DraftPresentation Create(Action<DraftPresentation> configure)
    {
        var draft = new DraftPresentation();
        configure(draft);
        return draft;
    }

    /// <inheritdoc />
    public ISlide Slide(int number) => this.Slides[number - 1];

    /// <summary>
    ///     Gets Slide Master by number.
    /// </summary>
    public ISlideMaster SlideMaster(int number) => this.SlideMasters[number - 1];

    /// <inheritdoc />
    public void Save()
    {
        // Materialize initial template slide if SlideIdList is empty but slide parts exist
        this.EnsureInitialSlideId();
        this.presDocument.PresentationPart!.Presentation.Save();
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
        // Materialize initial template slide if SlideIdList is empty but slide parts exist
        this.EnsureInitialSlideId();
        this.presDocument.PresentationPart!.Presentation.Save();

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

    private void EnsureInitialSlideId()
    {
        var presentationPart = this.presDocument.PresentationPart!;
        var presentation = presentationPart.Presentation;
        presentation.SlideIdList ??= new P.SlideIdList();
        var existingIds = presentation.SlideIdList.OfType<P.SlideId>().Select(s => s.RelationshipId!).ToHashSet();
        uint nextIdVal = presentation.SlideIdList.OfType<P.SlideId>().Any()
            ? presentation.SlideIdList.OfType<P.SlideId>().Max(s => s.Id!.Value) + 1u
            : 256u;

        // Ensure all slide parts are represented in SlideIdList
        foreach (var slidePart in presentationPart.SlideParts)
        {
            var relId = presentationPart.GetIdOfPart(slidePart);
            if (!existingIds.Contains(relId))
            {
                presentation.SlideIdList.Append(new P.SlideId { Id = nextIdVal++, RelationshipId = relId });
            }
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

    #region Fluent API

    /// <summary>
    ///     Represents a draft for building a presentation with a fluent API.
    /// </summary>
    public sealed class DraftPresentation
    {
        private readonly List<Action<Presentation>> actions = [];

        /// <summary>
        ///     Configures a slide within the presentation draft.
        ///     For a new presentation this targets the first slide.
        /// </summary>
        public DraftPresentation Slide(Action<DraftSlide> configure)
        {
            var slideDraft = new DraftSlide();
            configure(slideDraft);
            this.actions.Add(p => slideDraft.ApplyTo(p));
            return this;
        }

        /// <summary>
        ///     Generates a new presentation applying the configured actions.
        /// </summary>
        public Presentation Generate()
        {
            var presentation = new Presentation();
            foreach (var action in this.actions)
            {
                action(presentation);
            }

            return presentation;
        }
        
        internal void ApplyTo(Presentation presentation)
        {
            foreach (var action in this.actions)
            {
                action(presentation);
            }
        }
    }

    /// <summary>
    ///     Represents a draft for building a slide.
    /// </summary>
    public sealed class DraftSlide
    {
        private readonly List<Action<ISlide>> actions = [];
        
        /// <summary>
        ///     Adds a picture to the slide with the specified name and geometry in points.
        /// </summary>
        public DraftSlide Picture(string name, int x, int y, int width, int height, Stream image)
        {
            this.actions.Add(slide =>
            {
                // Add the picture
                slide.Shapes.AddPicture(image);

                // Modify the last added picture
                var picture = slide.Shapes.Last<IPicture>();
                picture.Name = name;
                picture.X = x;
                picture.Y = y;
                picture.Width = width;
                picture.Height = height;
            });

            return this;
        }

        /// <summary>
        ///     Configures a picture using a nested builder.
        /// </summary>
        public DraftSlide Picture(Action<PictureDraft> configure)
        {
            this.actions.Add(slide =>
            {
                var b = new PictureDraft();
                configure(b);
                slide.Shapes.AddPicture(b.ImageStream);
                var pic = slide.Shapes.Last<IPicture>();
                pic.Name = b.DraftName;
                pic.X = b.DraftX;
                pic.Y = b.DraftY;
                pic.Width = b.DraftWidth;
                pic.Height = b.DraftHeight;
                if (!string.IsNullOrEmpty(b.GeometryName))
                {
                    pic.GeometryType = (Geometry)Enum.Parse(typeof(Geometry), b.GeometryName.Replace(" ", string.Empty));
                }
            });

            return this;
        }
        
        /// <summary>
        ///     Adds a text box (auto shape) and sets its content.
        /// </summary>
        public DraftSlide TextBox(string name, int x, int y, int width, int height, string content)
        {
            this.actions.Add(slide =>
            {
                slide.Shapes.AddShape(x, y, width, height, Geometry.Rectangle, content);
                var addedShape = slide.Shapes.Last<IShape>();
                addedShape.Name = name;
            });

            return this;
        }

        /// <summary>
        ///     Configures a text box using a nested builder.
        /// </summary>
        public DraftSlide TextBox(Action<TextBoxDraft> configure)
        {
            this.actions.Add(slide =>
            {
                var builder = new TextBoxDraft();
                configure(builder);
                slide.Shapes.AddShape(builder.PosX, builder.PosY, builder.BoxWidth, builder.BoxHeight, Geometry.Rectangle);
                var addedShape = slide.Shapes.Last<IShape>();
                addedShape.Name = builder.BoxName;
                if (!string.IsNullOrEmpty(builder.Content))
                {
                    addedShape.TextBox!.SetText(builder.Content);
                }
            });

            return this;
        }

        /// <summary>
        ///     Adds a line shape.
        /// </summary>
        public DraftSlide Line(string name, int startPointX, int startPointY, int endPointX, int endPointY)
        {
            this.actions.Add(slide =>
            {
                slide.Shapes.AddLine(startPointX, startPointY, endPointX, endPointY);
                var line = slide.Shapes.Last<ILine>();
                line.Name = name;
            });

            return this;
        }
        
        /// <summary>
        ///     Adds a table with specified size.
        /// </summary>
        public DraftSlide Table(string name, int x, int y, int columnsCount, int rowsCount)
        {
            this.actions.Add(slide =>
            {
                slide.Shapes.AddTable(x, y, columnsCount, rowsCount);
                var table = slide.Shapes.Last<IShape>();
                table.Name = name;
            });

            return this;
        }
        
        internal void ApplyTo(Presentation presentation)
        {
            // Ensure there is at least one slide
            if (presentation.Slides.Count == 0)
            {
                // Ensure SlideIdList exists in the SDK presentation
                var sdkPres = presentation.presDocument.PresentationPart!.Presentation;
                sdkPres.SlideIdList ??= new P.SlideIdList();

                var blankLayout = presentation.SlideMasters[0].SlideLayouts.First(l => l.Name == "Blank");
                presentation.Slides.Add(blankLayout.Number);
            }

            // Target the first slide
            var slide = presentation.Slides[0];
            foreach (var action in this.actions)
            {
                action(slide);
            }
        }
    }

    public sealed class TextBoxDraft
    {
        internal string BoxName { get; private set; } = "TextBox";
        internal int PosX { get; private set; }
        internal int PosY { get; private set; }
        internal int BoxWidth { get; private set; } = 100;
        internal int BoxHeight { get; private set; } = 50;
        internal string? Content { get; private set; }

        public TextBoxDraft NameMethod(string name)
        {
            this.BoxName = name;
            return this;
        }

        public TextBoxDraft Name(string name) => this.NameMethod(name);

        public TextBoxDraft X(int x) { this.PosX = x; return this; }
        public TextBoxDraft Y(int y) { this.PosY = y; return this; }
        public TextBoxDraft Width(int width) { this.BoxWidth = width; return this; }
        public TextBoxDraft Height(int height) { this.BoxHeight = height; return this; }
        public TextBoxDraft Paragraph(string content) { this.Content = AppendParagraph(this.Content, content); return this; }

        private static string AppendParagraph(string? current, string next)
        {
            if (string.IsNullOrEmpty(current))
            {
                return next;
            }

            return current + Environment.NewLine + next;
        }
    }

    public sealed class PictureDraft
    {
        internal string DraftName { get; private set; } = "Picture";
        internal int DraftX { get; private set; }
        internal int DraftY { get; private set; }
        internal int DraftWidth { get; private set; } = 100;
        internal int DraftHeight { get; private set; } = 100;
        internal Stream ImageStream { get; private set; } = new MemoryStream();
        internal string? GeometryName { get; private set; }

        public PictureDraft NameMethod(string name) { this.DraftName = name; return this; }
        public PictureDraft Name(string name) => this.NameMethod(name);
        public PictureDraft X(int x) { this.DraftX = x; return this; }
        public PictureDraft Y(int y) { this.DraftY = y; return this; }
        public PictureDraft Width(int width) { this.DraftWidth = width; return this; }
        public PictureDraft Height(int height) { this.DraftHeight = height; return this; }
        public PictureDraft Image(Stream image) { this.ImageStream = image; return this; }
        public PictureDraft GeometryType(string geometry) { this.GeometryName = geometry; return this; }
    }

    #endregion
}