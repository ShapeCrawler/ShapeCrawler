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
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

#if DEBUG
using System.Threading.Tasks;
#endif

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a slide.
/// </summary>
public interface ISlide
{
    /// <summary>
    ///     Gets or sets custom data. Returns <see langword="null"/> if custom data is not presented.
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
    ///     Gets table by name.
    /// </summary>
    ITable Table(string name);

    /// <summary>
    ///     Gets picture by name.
    /// </summary>
    IPicture Picture(string picture);

    /// <summary>
    ///     Adds specified lines to the slide notes.
    /// </summary>
    void AddNotes(IEnumerable<string> lines);

    /// <summary>
    ///     Gets shape by name.
    /// </summary>
    /// <param name="name">Shape name.</param>
    /// <returns> An instance of <see cref="IShape"/>.</returns>
    IShape Shape(string name);

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
    ///     Gets chart by name.
    /// </summary>
    IChart Chart(string name);
    
    /// <summary>
    ///     Gets chart by ID.
    /// </summary>
    IChart Chart(int id);
}

internal sealed class Slide : ISlide
{
    private readonly SlidePart slidePart;
    private CustomXmlPart? customDataCustomXmlPart;
    private IShapeFill? fill;

    internal Slide(
        SlidePart slidePart,
        ISlideLayout slideLayout,
        MediaCollection mediaCollection)
    {
        this.slidePart = slidePart;
        this.customDataCustomXmlPart = this.GetCustomXmlPart();
        this.SlideLayout = slideLayout;
        this.Shapes = new SlideShapeCollection(new ShapeCollection(slidePart), this.slidePart, mediaCollection);
    }

    public ISlideLayout SlideLayout { get; }

    public ISlideShapeCollection Shapes { get; }

    public int Number
    {
        get => this.ParseNumber();
        set
        {
            if (this.Number == value)
            {
                return;
            }

            var currentIndex = this.Number - 1;
            var newIndex = value - 1;
            var presDocument = (PresentationDocument)this.slidePart.OpenXmlPackage;
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

            // Insert the source slide at its new position after the target slide.
            slideIdList.InsertAfter(sourceSlide, targetSlide);

            // Save the modified presentation.
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
                var pcSld = this.slidePart.Slide.CommonSlideData
                            ?? this.slidePart.Slide.AppendChild<DocumentFormat.OpenXml.Presentation.CommonSlideData>(
                                new());

                // Background element needs to be first, else it gets ignored.
                var pBg = pcSld.GetFirstChild<DocumentFormat.OpenXml.Presentation.Background>()
                          ?? pcSld.InsertAt<DocumentFormat.OpenXml.Presentation.Background>(new(), 0);

                var pBgPr = pBg.GetFirstChild<DocumentFormat.OpenXml.Presentation.BackgroundProperties>()
                            ?? pBg.AppendChild<DocumentFormat.OpenXml.Presentation.BackgroundProperties>(new());

                this.fill = new ShapeFill(pBgPr);
            }

            return this.fill!;
        }
    }

    public bool Hidden() => this.slidePart.Slide.Show is not null && !this.slidePart.Slide.Show.Value;

    public void Hide()
    {
        if (this.slidePart.Slide.Show is null)
        {
            var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
            this.slidePart.Slide.SetAttribute(showAttribute);
        }
        else
        {
            this.slidePart.Slide.Show = false;
        }
    }

    public ITable Table(string name) => this.Shapes.Shape<ITable>(name);

    public IPicture Picture(string picture) => this.Shapes.Shape<IPicture>(picture);

    public IShape Shape(string name) => this.Shapes.Shape<IShape>(name);

    public T Shape<T>(string name)
        where T : IShape
        => this.Shapes.Shape<T>(name);

    public void Remove()
    {
        // TODO: slide layout and master of removed slide also should be deleted if they are unused
        var presDocument = (PresentationDocument)this.slidePart.OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pPresentation = presDocument.PresentationPart!.Presentation;
        var slideIdList = pPresentation.SlideIdList!;
        var removingPSlideId = (P.SlideId)slideIdList.ChildElements[this.Number - 1];
        var sectionList = pPresentation.PresentationExtensionList?.Descendants<P14.SectionList>().FirstOrDefault();
        var removingSectionSlideIdListEntry = sectionList?.Descendants<P14.SectionSlideIdListEntry>()
            .FirstOrDefault(s => s.Id! == removingPSlideId.Id!);
        removingSectionSlideIdListEntry?.Remove();
        slideIdList.RemoveChild(removingPSlideId);
        pPresentation.Save();

        var removingSlideIdRelationshipId = removingPSlideId.RelationshipId!;
        new SCPPresentation(pPresentation).RemoveSlideIdFromCustomShow(removingSlideIdRelationshipId.Value!);

        var removingSlidePart = (SlidePart)presPart.GetPartById(removingSlideIdRelationshipId!);
        presPart.DeletePart(removingSlidePart);

        presPart.Presentation.Save();
    }

    public IChart Chart(string name) => this.Shapes.Shape<IChart>(name);

    public IChart Chart(int id) => this.Shapes.GetById<IChart>(id);

    public IList<ITextBox> GetTextBoxes()
    {
        var textBoxes = this.Shapes
            .Where(shape => shape.TextBox is not null)
            .Select(shape => shape.TextBox!)
            .ToList();

        var tableTextboxes = this.Shapes.OfType<ITable>().SelectMany(table => table.Rows.SelectMany(row => row.Cells))
            .Where(cell => cell.TextBox is not null).Select(cell => cell.TextBox);
        textBoxes.AddRange(tableTextboxes);

        var groupShapes = this.Shapes.OfType<Group>().ToList();
        foreach (var groupShape in groupShapes)
        {
            this.AddGroupTextBoxes(groupShape, textBoxes);
        }

        return textBoxes;
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

    internal PresentationDocument SdkPresentationDocument() => (PresentationDocument)this.slidePart.OpenXmlPackage;

    private void AddGroupTextBoxes(IGroup groupShape, List<ITextBox> textBoxes)
    {
        foreach (var shape in groupShape.Shapes)
        {
            if (shape is IGroup group)
            {
                this.AddGroupTextBoxes(group, textBoxes);
            }
            else if(shape.TextBox is not null)
            {
                textBoxes.Add(shape.TextBox);
            }
        }
    }

    private ITextBox? GetNotes()
    {
        var notes = this.slidePart.NotesSlidePart;

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
        var rid = new SCOpenXmlPart(this.slidePart).GetNextRelationshipId();
        var notesSlidePart1 = this.slidePart.AddNewPart<NotesSlidePart>(rid);
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

    private int ParseNumber()
    {
        var sdkPresentationDocument = (PresentationDocument)this.slidePart.OpenXmlPackage;
        var presentationPart = sdkPresentationDocument.PresentationPart!;
        var currentSlidePartId = presentationPart.GetIdOfPart(this.slidePart);
        var slideIdList =
            presentationPart.Presentation.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
        for (var i = 0; i < slideIdList.Count; i++)
        {
            if (slideIdList[i].RelationshipId == currentSlidePartId)
            {
                return i + 1;
            }
        }

        throw new SCException("An error occurred while parsing slide number.");
    }

    private string? GetCustomData()
    {
        if (this.customDataCustomXmlPart == null)
        {
            return null;
        }

        var customXmlPartStream = this.customDataCustomXmlPart.GetStream();
        using var customXmlStreamReader = new StreamReader(customXmlPartStream);
        var raw = customXmlStreamReader.ReadToEnd();
        return raw[3..];
    }

    private void SetCustomData(string? value)
    {
        Stream customXmlPartStream;
        if (this.customDataCustomXmlPart == null)
        {
            var newSlideCustomXmlPart = this.slidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPartStream = newSlideCustomXmlPart.GetStream();
            this.customDataCustomXmlPart = newSlideCustomXmlPart;
        }
        else
        {
            customXmlPartStream = this.customDataCustomXmlPart.GetStream();
        }

        using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
        customXmlStreamReader.Write($"ctd{value}");
    }

    private CustomXmlPart? GetCustomXmlPart()
    {
        foreach (var customXmlPart in this.slidePart.CustomXmlParts)
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