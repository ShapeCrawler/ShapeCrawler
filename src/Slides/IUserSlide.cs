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
using ShapeCrawler.Units;
using SkiaSharp;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

#if DEBUG
using System.Threading.Tasks;
#endif

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a slide.
/// </summary>
public interface IUserSlide
{
    /// <summary>
    ///     Gets or sets custom data. Returns <see langword="null"/> if the custom data is not presented.
    /// </summary>
    string? CustomData { get; set; }

    /// <summary>
    ///     Gets slide layout.
    /// </summary>
    ILayoutSlide LayoutSlide { get; }

    /// <summary>
    ///     Gets or sets slide number.
    /// </summary>
    int Number { get; set; }

    /// <summary>
    ///     Gets the shape collection.
    /// </summary>
    IUserSlideShapeCollection Shapes { get; }

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
    ///     Saves the slide image to the specified stream.
    /// </summary>
    void SaveImageTo(Stream stream);

#if DEBUG
    /// <summary>
    ///     Saves the slide image to the specified file.
    /// </summary>
    void SaveImageTo(string file);
#endif

    /// <summary>
    ///     Gets a copy of the underlying parent <see cref="PresentationPart"/>.
    /// </summary>
    // ReSharper disable once InconsistentNaming
    PresentationPart GetSdkPresentationPart();

    /// <summary>
    ///     Gets the first shape in the slide.
    /// </summary>
    /// <typeparam name="T">Shape type.</typeparam>
    T First<T>();
}

internal class UserSlide(ILayoutSlide layoutSlide, UserSlideShapeCollection shapes, SlidePart slidePart) : IUserSlide
{
    private IShapeFill? fill;

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
            var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
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
                var pcSld = slidePart.Slide.CommonSlideData
                            ?? slidePart.Slide.AppendChild<CommonSlideData>(
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

    public bool Hidden() => slidePart.Slide.Show is not null && !slidePart.Slide.Show.Value;

    public void Hide()
    {
        if (slidePart.Slide.Show is null)
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

    public void SaveImageTo(string file)
    {
        using var fileStream = File.Create(file);
        this.SaveImageTo(fileStream);
    }

    public PresentationPart GetSdkPresentationPart()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;

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

    public void Remove()
    {
        var presDocument = (PresentationDocument)slidePart.OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pPresentation = presDocument.PresentationPart!.Presentation;
        var slideIdList = pPresentation.SlideIdList!;

        // Find the exact SlideId corresponding to this slide
        var slideIdRelationship = presPart.GetIdOfPart(slidePart);
        var removingPSlideId = slideIdList.Elements<P.SlideId>()
                                   .FirstOrDefault(slideId => slideId.RelationshipId!.Value == slideIdRelationship) ??
                               throw new SCException("Could not find slide ID in presentation.");

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

    public void SaveImageTo(Stream stream)
    {
        var presPart = this.GetSdkPresentationPart();
        var pSlideSize = presPart.Presentation.SlideSize!;
        var width = new Emus(pSlideSize.Cx!.Value).AsPixels();
        var height = new Emus(pSlideSize.Cy!.Value).AsPixels();

        using var surface = SKSurface.Create(new SKImageInfo((int)width, (int)height));
        var canvas = surface.Canvas;

        this.RenderBackground(canvas);
        shapes.Render(canvas);

        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        data.SaveTo(stream);

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
        var rid = new SCOpenXmlPart(slidePart).NextRelationshipId();
        var notesSlidePart1 = slidePart.AddNewPart<NotesSlidePart>(rid);
        var notesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                        {
                            Id = (UInt32Value)1U, Name = string.Empty
                        },
                        new DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new DocumentFormat.OpenXml.Presentation.Shape(
                        new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                            new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties
                            {
                                Id = (UInt32Value)2U, Name = "Notes Placeholder 2"
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

    private SKColor GetSkColor()
    {
        var hex = this.Fill.Color!.TrimStart('#');

        // Validate hex length before parsing
        if (hex.Length != 6 && hex.Length != 8)
        {
            return SKColors.White; // used by the PowerPoint application as the default background color
        }

        return new Color(hex).AsSkColor();
    }

    private void RenderBackground(SKCanvas canvas)
    {
        var slideFill = this.Fill;
        if (slideFill is { Type: FillType.Solid, Color: not null })
        {
            var skColor = this.GetSkColor();
            canvas.Clear(skColor);
        }
        else if (slideFill is { Type: FillType.Picture, Picture: not null })
        {
            var bytes = slideFill.Picture.AsByteArray();
            using var stream = new MemoryStream(bytes);
            using var bitmap = SKBitmap.Decode(stream);
            var destRect = new SKRect(0, 0, canvas.DeviceClipBounds.Width, canvas.DeviceClipBounds.Height);
            canvas.DrawBitmap(bitmap, destRect);
        }
        else
        {
            // Default to white for unsupported backgrounds
            canvas.Clear(SKColors.White);
        }
    }
}