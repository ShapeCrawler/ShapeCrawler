using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Drawing;
using ShapeCrawler.Presentations;
using ShapeCrawler.Shapes;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;

#if DEBUG
using System.Threading.Tasks;
#endif

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a slide.
/// </summary>
public interface ISlide
{
    /// <summary>
    ///     Gets or sets custom data. Returns <see langword="null"/> if the custom data is not presented.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets slide layout.
    /// </summary>
    ISlideLayout SlideLayout { get; }

    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }

    /// <summary>
    ///     Gets the shape collection.
    /// </summary>
    ISlideShapeCollection Shapes { get; }

    /// <summary>
    ///     Gets the slide notes.
    /// </summary>
    ITextBox? Notes { get; }

    /// <summary>
    ///     Gets the slide fill.
    /// </summary>
    IShapeFill Fill { get; }

    /// <summary>
    ///     Gets all slide text boxes.
    /// </summary>
    public IList<ITextBox> GetTextBoxes();

    /// <summary>
    ///     Hides slide.
    /// </summary>
    void Hide();

    /// <summary>
    ///     Gets a value indicating whether the slide is hidden.
    /// </summary>
    bool Hidden();

    /// <summary>
    ///     Adds specified lines to the slide notes.
    /// </summary>
    void AddNotes(IEnumerable<string> lines);

    /// <summary>
    ///     Gets element by name.
    /// </summary>
    /// <param name="name">element name.</param>
    IShape Shape(string name);
    
    /// <summary>
    ///     Gets element by ID.
    /// </summary>
    IShape Shape(int id);

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T Shape<T>(string name)
        where T : IShape;

    /// <summary>
    ///     Removes the slide.
    /// </summary>
    void Remove();

    /// <summary>
    ///     Saves the slide as an image.
    /// </summary>
    void SaveAsImage(Stream stream);

    /// <summary>
    ///     Gets a copy of the underlying parent <see cref="PresentationPart"/>.
    /// </summary>
    // ReSharper disable once InconsistentNaming
    PresentationPart GetSDKPresentationPart();

    /// <summary>
    ///     Gets the first shape in the slide.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T First<T>();
}

internal abstract class Slide : ISlide
{
    protected readonly SlidePart SlidePart;
    private IShapeFill? fill;

    private protected Slide(ISlideLayout slideLayout, ISlideShapeCollection shapes, SlidePart slidePart)
    {
        this.SlideLayout = slideLayout;
        this.Shapes = shapes;
        this.SlidePart = slidePart;
    }

    public ISlideLayout SlideLayout { get; }

    public ISlideShapeCollection Shapes { get; }

    public int Number
    {
        get
        {
            var presDocument = (PresentationDocument)this.SlidePart.OpenXmlPackage;
            var presPart = presDocument.PresentationPart!;
            var currentSlidePartId = presPart.GetIdOfPart(this.SlidePart);
            var slideIdList =
                presPart.Presentation.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
            for (var i = 0; i < slideIdList.Count; i++)
            {
                if (slideIdList[i].RelationshipId == currentSlidePartId)
                {
                    return i + 1;
                }
            }

            throw new SCException("An error occurred while parsing slide number.");
        }

        set
        {
            if (this.Number == value)
            {
                throw new SCException("Slide number is already set to the specified value.");
            }

            var currentIndex = this.Number - 1;
            var newIndex = value - 1;
            var presDocument = (PresentationDocument)this.SlidePart.OpenXmlPackage;
            if (newIndex < 0 || newIndex >= presDocument.PresentationPart!.SlideParts.Count())
            {
                throw new SCException("Slide number is out of range.");
            }

            var presentationPart = presDocument.PresentationPart!;
            var presentation = presentationPart.Presentation;
            var slideIdList = presentation.SlideIdList!;

            // Get the slide ID of the source slide.
            var sourceSlide = (SlideId)slideIdList.ChildElements[currentIndex];

            SlideId? targetSlide;

            // Identify the position of the target slide after which to move the source slide
            if (newIndex == 0)
            {
                targetSlide = null;
            }
            else if (currentIndex < newIndex)
            {
                targetSlide = (SlideId)slideIdList.ChildElements[newIndex];
            }
            else
            {
                targetSlide = (SlideId)slideIdList.ChildElements[newIndex - 1];
            }

            // Remove the source slide from its current position.
            sourceSlide.Remove();
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            presentation.Save();
        }
    }

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value);
    }

    public ITextBox? Notes => this.GetNotes();

    public IShapeFill Fill
    {
        get
        {
            if (this.fill is null)
            {
                var pcSld = this.SlidePart.Slide.CommonSlideData
                            ?? this.SlidePart.Slide.AppendChild<CommonSlideData>(
                                new());

                // Background element needs to be first, else it gets ignored.
                var pBg = pcSld.GetFirstChild<Background>()
                          ?? pcSld.InsertAt<Background>(new(), 0);

                var pBgPr = pBg.GetFirstChild<P.BackgroundProperties>()
                            ?? pBg.AppendChild<BackgroundProperties>(new());

                this.fill = new ShapeFill(pBgPr);
            }

            return this.fill!;
        }
    }

    public bool Hidden() => this.SlidePart.Slide.Show is not null && !this.SlidePart.Slide.Show.Value;

    public void Hide()
    {
        if (this.SlidePart.Slide.Show is null)
        {
            var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
            this.SlidePart.Slide.SetAttribute(showAttribute);
        }
        else
        {
            this.SlidePart.Slide.Show = false;
        }
    }

    public IShape Shape(string name) => this.Shapes.Shape<IShape>(name);

    public IShape Shape(int id) => this.Shapes.GetById<IShape>(id);

    public T Shape<T>(string name)
        where T : IShape
        => this.Shapes.Shape<T>(name);

    public PresentationPart GetSDKPresentationPart()
    {
        var presDocument = (PresentationDocument)this.SlidePart.OpenXmlPackage;

        return presDocument.Clone().PresentationPart!;
    }

    public T First<T>() => (T)this.Shapes.First(shape => shape is T);

    public IList<ITextBox> GetTextBoxes()
    {
        var collectedTextBoxes = new List<ITextBox>();

        foreach (var shape in this.Shapes)
        {
            this.CollectTextBoxes(shape, collectedTextBoxes);
        }

        return collectedTextBoxes;
    }

    /// <inheritdoc/>
    public void AddNotes(IEnumerable<string> lines)
    {
        var notes = this.Notes;
        if (notes is null)
        {
            this.AddNotesSlide(lines);
        }
        else
        {
            var paragraphs = notes.Paragraphs;
            foreach (var line in lines)
            {
                paragraphs.Add();
                paragraphs[paragraphs.Count - 1].Text = line;
            }
        }
    }

    public abstract void Remove(); 

    public void SaveAsImage(Stream stream)
    {
        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        var slideImage = new SlideImage(this);
        slideImage.Save(stream, SKEncodedImageFormat.Png);
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }
    }

    private void CollectTextBoxes(IShape shape, List<ITextBox> buffer)
    {
        if (shape.TextBox is not null)
        {
            buffer.Add(shape.TextBox);
        }

        if (shape.Table is not null)
        {
            foreach (var cell in shape.Table.Rows.SelectMany(row => row.Cells))
            {
                buffer.Add(cell.TextBox);
            }
        }

        if (shape.GroupedShapes is not null)
        {
            foreach (var innerShape in shape.GroupedShapes)
            {
                this.CollectTextBoxes(innerShape, buffer);
            }
        }
    }

    private ITextBox? GetNotes()
    {
        var notes = this.SlidePart.NotesSlidePart;

        if (notes is null)
        {
            return null;
        }

        var shapes = new ShapeCollection(notes);
        var notesPlaceholder = shapes
            .FirstOrDefault(shape =>
                shape is { PlaceholderType: not null, TextBox: not null, PlaceholderType: PlaceholderType.Text });
        return notesPlaceholder?.TextBox;
    }

    private void AddNotesSlide(IEnumerable<string> lines)
    {
        // Build up the children of the text body element
        var textBodyChildren = new List<OpenXmlElement>() { new BodyProperties(), new ListStyle() };

        // Add in the text lines
        textBodyChildren.AddRange(
            lines
                .Select(line => new DocumentFormat.OpenXml.Drawing.Paragraph(
                    new ParagraphProperties(),
                    new Run(
                        new RunProperties(),
                        new DocumentFormat.OpenXml.Drawing.Text(line)),
                    new EndParagraphRunProperties())));

        // Always add at least one paragraph, even if empty
        if (!lines.Any())
        {
            textBodyChildren.Add(
                new DocumentFormat.OpenXml.Drawing.Paragraph(
                    new EndParagraphRunProperties()));
        }

        // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides
        var rid = new SCOpenXmlPart(this.SlidePart).NextRelationshipId();
        var notesSlidePart1 = this.SlidePart.AddNewPart<NotesSlidePart>(rid);
        var notesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties()
                        {
                            Id = (UInt32Value)1U, Name = string.Empty
                        },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new DocumentFormat.OpenXml.Presentation.Shape(
                        new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties()
                            {
                                Id = (UInt32Value)2U, Name = "Notes Placeholder 2"
                            },
                            new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(
                                new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape() { Type = PlaceholderValues.Body })),
                        new DocumentFormat.OpenXml.Presentation.ShapeProperties(),
                        new DocumentFormat.OpenXml.Presentation.TextBody(
                            textBodyChildren)))),
            new ColorMapOverride(new MasterColorMapping()));
        notesSlidePart1.NotesSlide = notesSlide;
    }

    private string? GetCustomData()
    {
        var getCustomXmlPart = this.GetCustomXmlPartOrNull();
        if (getCustomXmlPart == null)
        {
            return null;
        }

        var customXmlPartStream = getCustomXmlPart.GetStream();
        using var customXmlStreamReader = new StreamReader(customXmlPartStream);
        var raw = customXmlStreamReader.ReadToEnd();
        return raw[3..];
    }

    private void SetCustomData(string? value)
    {
        var getCustomXmlPart = this.GetCustomXmlPartOrNull();
        Stream customXmlPartStream;
        if (getCustomXmlPart == null)
        {
            var newSlideCustomXmlPart = this.SlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPartStream = newSlideCustomXmlPart.GetStream();
        }
        else
        {
            customXmlPartStream = getCustomXmlPart.GetStream();
        }

        using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
        customXmlStreamReader.Write($"ctd{value}");
    }

    private CustomXmlPart? GetCustomXmlPartOrNull()
    {
        foreach (var customXmlPart in this.SlidePart.CustomXmlParts)
        {
            using var customXmlPartStream = new StreamReader(customXmlPart.GetStream());
            var customXmlPartText = customXmlPartStream.ReadToEnd();
            if (customXmlPartText.StartsWith(
                    "ctd",
                    StringComparison.CurrentCulture))
            {
                return customXmlPart;
            }
        }

        return null;
    }
}