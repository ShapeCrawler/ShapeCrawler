using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
#if NETSTANDARD2_0
using System.Collections.Generic;
using ShapeCrawler.Extensions;
#endif
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Presentations;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <inheritdoc />
public sealed class Presentation : IPresentation
{
    internal readonly PresentationDocument PresDocument;
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

        this.PresDocument = PresentationDocument.Open(this.presStream, true);
        this.slideSize = new SlideSize(this.PresDocument.PresentationPart!.Presentation!.SlideSize!);
        this.MasterSlides = new MasterSlideCollection(this.PresDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollection(this.PresDocument);
        this.Slides = new UpdatedSlideCollection(
            new UserSlideCollection(this.PresDocument.PresentationPart.SlideParts),
            this.PresDocument.PresentationPart);
        this.Footer = new Footer(new UpdatedSlideCollection(
            new UserSlideCollection(this.PresDocument.PresentationPart.SlideParts), this.PresDocument.PresentationPart));
        this.Properties =
            this.PresDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(this.PresDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
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

        this.PresDocument = PresentationDocument.Open(this.presStream, true);
        this.slideSize = new SlideSize(this.PresDocument.PresentationPart!.Presentation!.SlideSize!);
        this.MasterSlides = new MasterSlideCollection(this.PresDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollection(this.PresDocument);
        this.Slides = new UpdatedSlideCollection(
            new UserSlideCollection(this.PresDocument.PresentationPart.SlideParts),
            this.PresDocument.PresentationPart);
        this.Footer = new Footer(new UpdatedSlideCollection(
            new UserSlideCollection(this.PresDocument.PresentationPart.SlideParts), this.PresDocument.PresentationPart));
        this.Properties =
            this.PresDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(this.PresDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
    }

    /// <summary>
    ///     Creates a new presentation.
    /// </summary>
    public Presentation()
    {
        this.presStream = new AssetCollection(Assembly.GetExecutingAssembly()).StreamOf("new presentation.pptx");

        this.PresDocument = PresentationDocument.Open(this.presStream, true);
        this.slideSize = new SlideSize(this.PresDocument.PresentationPart!.Presentation!.SlideSize!);
        this.MasterSlides = new MasterSlideCollection(this.PresDocument.PresentationPart!.SlideMasterParts);
        this.Sections = new SectionCollection(this.PresDocument);
        this.Slides = new UpdatedSlideCollection(
            new UserSlideCollection(this.PresDocument.PresentationPart.SlideParts),
            this.PresDocument.PresentationPart);
        this.Footer = new Footer(new UpdatedSlideCollection(
            new UserSlideCollection(this.PresDocument.PresentationPart.SlideParts), this.PresDocument.PresentationPart));
        this.Properties =
            this.PresDocument.CoreFilePropertiesPart != null
                ? new PresentationProperties(this.PresDocument.CoreFilePropertiesPart.OpenXmlPackage.PackageProperties)
                : new PresentationProperties(new DefaultPackageProperties());
        this.Properties.Modified = SCSettings.TimeProvider.UtcNow;
    }

    /// <summary>
    ///     Creates a new presentation using fluent configuration.
    /// </summary>
    public Presentation(Action<DraftPresentation> configure)
        : this()
    {
        var draft = new DraftPresentation(this);
        configure(draft);
        draft.ApplyTo(this);
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
    public IMasterSlideCollection MasterSlides { get; }

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
    public IUserSlide Slide(int number)
    {
        if (number < 0)
        {
            throw new SCException($"Specified slide number is must {number} be more than zero.");
        }

        return number > this.Slides.Count ? throw new SCException($"Specified slide number {number} exceeds the number of slides {this.Slides.Count} in the presentation.") : this.Slides[number - 1];
    }

    /// <summary>
    ///     Gets Slide Master by number.
    /// </summary>
    public IMasterSlide SlideMaster(int number) => this.MasterSlides[number - 1];

    /// <inheritdoc />
    public void Save()
    {
        // Materialize initial template slide if SlideIdList is empty but slide parts exist
        this.EnsureInitialSlideId();
        this.PresDocument.PresentationPart!.Presentation!.Save();
        this.PresDocument.Save();
        if (this.inputPresStream is not null)
        {
            this.PresDocument.Clone(this.inputPresStream);
        }
        else if (this.inputPresFile is not null)
        {
            var savedPres = this.PresDocument.Clone(this.inputPresFile);
            savedPres.Dispose();
        }
    }

    /// <inheritdoc />
    public void Save(Stream stream)
    {
        this.Properties.Modified = SCSettings.TimeProvider.UtcNow;
        this.EnsureInitialSlideId();
        this.PresDocument.PresentationPart!.Presentation!.Save();

        if (stream is FileStream fileStream)
        {
            var mStream = new MemoryStream();
            this.PresDocument.Clone(mStream);
            mStream.Position = 0;
            mStream.CopyTo(fileStream);
        }
        else
        {
            this.PresDocument.Clone(stream);
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
            var textShapes = slide.Shapes
                .Select(shape => new { shape, shapeText = shape.TextBox })
                .Where(x => x.shapeText is not null
                            && x.shapeText.Text != string.Empty
                            && x.shape.PlaceholderType != PlaceholderType.SlideNumber);

            var titleShape = textShapes.FirstOrDefault(x =>
                x.shape.Name.StartsWith("Title", StringComparison.OrdinalIgnoreCase));
            if (titleShape != null)
            {
                markdown.AppendLine($"## {titleShape.shapeText!.Text}");
            }

            foreach (var nonTitleShape in textShapes
                         .Where(x => !x.shape.Name.StartsWith("Title", StringComparison.OrdinalIgnoreCase)))
            {
                markdown.AppendLine(nonTitleShape.shapeText!.Text);
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
    public PresentationDocument GetSdkPresentationDocument() => this.PresDocument.Clone();

    /// <summary>
    ///     Releases all resources used by the presentation.
    /// </summary>
    public void Dispose() => this.PresDocument.Dispose();

    private void EnsureInitialSlideId()
    {
        var presentationPart = this.PresDocument.PresentationPart!;
        var presentation = presentationPart.Presentation!;
        presentation.SlideIdList ??= new P.SlideIdList();
#if NETSTANDARD2_0
        var existingIds = new HashSet<string>(
            presentation.SlideIdList
                .OfType<P.SlideId>()
                .Select(s => (string)s.RelationshipId!));
#else
        var existingIds = presentation.SlideIdList
            .OfType<P.SlideId>()
            .Select(s => (string)s.RelationshipId!)
            .ToHashSet();
#endif
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

    #region Fluent API

    #endregion
}