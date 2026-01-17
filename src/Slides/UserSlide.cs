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
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Slides;

/// <inheritdoc/>
internal abstract class UserSlide(ILayoutSlide layoutSlide, UserSlideShapeCollection shapes, SlidePart slidePart) : IUserSlide
{
    public ILayoutSlide LayoutSlide => layoutSlide;

    public IUserSlideShapeCollection Shapes => shapes;

    public int Number
    {
        get
        {
            var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
            var presPart = presDocument.PresentationPart!;
            var currentSlidePartId = presPart.GetIdOfPart(slidePart);
            var slideIdList =
                presPart.Presentation!.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
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
            var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
            if (newIndex < 0 || newIndex >= presDocument.PresentationPart!.SlideParts.Count())
            {
                throw new SCException("Slide number is out of range.");
            }

            var presentationPart = presDocument.PresentationPart!;
            var presentation = presentationPart.Presentation;
            var slideIdList = presentation!.SlideIdList!;

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
            if (field is not null)
            {
                return field!;
            }

            var pcSld = slidePart.Slide!.CommonSlideData
                        ?? slidePart.Slide!.AppendChild(
                            new CommonSlideData());

            // Background element needs to be first, else it gets ignored.
            var pBg = pcSld.GetFirstChild<Background>()
                      ?? pcSld.InsertAt<Background>(new(), 0);

            var pBgPr = pBg.GetFirstChild<BackgroundProperties>();
            if (pBgPr is null)
            {
                // PowerPoint always keeps background properties schema-valid.
                // If we create an empty p:bgPr, Open XML validation fails because it must contain a fill element.
                pBgPr = new BackgroundProperties(new NoFill());
                pBg.AppendChild(pBgPr);
            }
            else
            {
                var hasFill =
                    pBgPr.GetFirstChild<A.BlipFill>() is not null
                    || pBgPr.GetFirstChild<GradientFill>() is not null
                    || pBgPr.GetFirstChild<NoFill>() is not null;
                hasFill = hasFill
                          || pBgPr.GetFirstChild<PatternFill>() is not null
                          || pBgPr.GetFirstChild<SolidFill>() is not null;
                if (!hasFill)
                {
                    // Keep schema-valid even if p:bgPr was previously created empty.
                    pBgPr.InsertAt(new NoFill(), 0);
                }
            }

            field = new ShapeFill(pBgPr);

            return field!;
        }
    }

    public bool Hidden() => slidePart.Slide!.Show is not null && !slidePart.Slide!.Show.Value;

    public void Hide()
    {
        if (slidePart.Slide!.Show is null)
        {
            var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
            slidePart.Slide.SetAttribute(showAttribute);
        }
        else
        {
            slidePart.Slide.Show = false;
        }
    }

    public IShape Shape(string name) => this.Shapes.Shape<IShape>(name);

    public IShape Shape(int id) => this.Shapes.GetById<IShape>(id);

    public T Shape<T>(string name)
        where T : IShape
        => this.Shapes.Shape<T>(name);

    /// <inheritdoc/>
    public abstract void SaveImageTo(string file);

    /// <inheritdoc/>
    public abstract void SaveImageTo(Stream stream);

    public PresentationPart GetSdkPresentationPart()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;

        return presDocument.Clone().PresentationPart!;
    }

    public T First<T>() => (T)this.Shapes.First(shape => shape is T);

    public IList<ITextBox> GetTexts()
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

    public void Remove()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pPresentation = presDocument.PresentationPart!.Presentation!;
        var slideIdList = pPresentation.SlideIdList!;

        // Find the exact SlideId corresponding to this slide
        var slideIdRelationship = presPart.GetIdOfPart(slidePart);
        var removingPSlideId = slideIdList.Elements<SlideId>()
                                   .FirstOrDefault(slideId => slideId.RelationshipId!.Value == slideIdRelationship) ??
                               throw new SCException("Could not find slide ID in presentation.");

        var sectionList = pPresentation.PresentationExtensionList?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.SectionList>().FirstOrDefault();
        var removingSectionSlideIdListEntry = sectionList?.Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.SectionSlideIdListEntry>()
            .FirstOrDefault(s => s.Id! == removingPSlideId.Id!);
        removingSectionSlideIdListEntry?.Remove();

        slideIdList.RemoveChild(removingPSlideId);
        pPresentation.Save();

        var removingSlideIdRelationshipId = removingPSlideId.RelationshipId!;
        new SCPPresentation(pPresentation).RemoveSlideIdFromCustomShow(removingSlideIdRelationshipId.Value!);

        var removingSlidePart = (SlidePart)presPart.GetPartById(removingSlideIdRelationshipId!);
        presPart.DeletePart(removingSlidePart);

        presPart.Presentation!.Save();
    }
    
    /// <summary>
    ///     Gets the underlying <see cref="SlidePart"/>.
    /// </summary>
    /// <returns>Slide part instance.</returns>
    internal SlidePart GetSdkSlidePart() => slidePart;

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
        var notesSlidePart = slidePart.NotesSlidePart;

        if (notesSlidePart is null)
        {
            return null;
        }

        var notesShapes = new ShapeCollection(notesSlidePart);
        var notesPlaceholder = notesShapes
            .FirstOrDefault(shape =>
                shape is { PlaceholderType: not null, TextBox: not null, PlaceholderType: PlaceholderType.Text });
        return notesPlaceholder?.TextBox;
    }

    private void AddNotesSlide(IEnumerable<string> lines)
    {
        // Build up the children of the text body element
        var textBodyChildren = new List<OpenXmlElement> { new BodyProperties(), new ListStyle() };

        // Add in the text lines
        textBodyChildren.AddRange(
            lines
                .Select(line => new A.Paragraph(
                    new ParagraphProperties(),
                    new Run(
                        new RunProperties(),
                        new A.Text(line)),
                    new EndParagraphRunProperties())));

        // Always add at least one paragraph, even if empty
        if (!lines.Any())
        {
            textBodyChildren.Add(
                new DocumentFormat.OpenXml.Drawing.Paragraph(
                    new EndParagraphRunProperties()));
        }

        // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides
        var rid = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var notesSlidePart1 = slidePart.AddNewPart<NotesSlidePart>(rid);
        var notesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                        {
                            Id = (UInt32Value)1U,
                            Name = string.Empty
                        },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new DocumentFormat.OpenXml.Presentation.Shape(
                        new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                            {
                                Id = (UInt32Value)2U,
                                Name = "Notes Placeholder 2"
                            },
                            new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(
                                new ShapeLocks { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(
                                new PlaceholderShape { Type = PlaceholderValues.Body })),
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
            var newSlideCustomXmlPart = slidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
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
        foreach (var customXmlPart in slidePart.CustomXmlParts)
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