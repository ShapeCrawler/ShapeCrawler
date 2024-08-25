using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shared;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler;

internal sealed class Slide : ISlide
{
    private readonly Lazy<SlideBgImage> backgroundImage;
    private readonly SlideSize slideSize;
    private Lazy<CustomXmlPart?> sdkCustomXmlPart;

    internal Slide(
        SlidePart sdkSlidePart,
        ISlideLayout slideLayout,
        SlideSize slideSize)
    {
        this.SDKSlidePart = sdkSlidePart;
        this.slideSize = slideSize;
        this.backgroundImage = new Lazy<SlideBgImage>(() =>
            new SlideBgImage(sdkSlidePart));
        this.sdkCustomXmlPart = new Lazy<CustomXmlPart?>(this.GetSldCustomXmlPart);
        this.SlideLayout = slideLayout;
        this.Shapes = new SlideShapes(this.SDKSlidePart, new ShapeCollection.Shapes(sdkSlidePart));
    }

    public ISlideLayout SlideLayout { get; }
    
    public SlidePart SDKSlidePart { get; }

    public ISlideShapes Shapes { get; }

    public int Number
    {
        get => this.ParseNumber();
        set => this.UpdateNumber(value);
    }

    public IImage? Background => this.backgroundImage.Value;

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value);
    }

    public ITextBox? Notes => this.GetNotes();

    public bool Hidden() => this.SDKSlidePart.Slide.Show is not null && !this.SDKSlidePart.Slide.Show.Value;

    public void Hide()
    {
        if (this.SDKSlidePart.Slide.Show is null)
        {
            var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
            this.SDKSlidePart.Slide.SetAttribute(showAttribute);
        }
        else
        {
            this.SDKSlidePart.Slide.Show = false;
        }
    }
    
    public ITable Table(string name) => this.Shapes.GetByName<ITable>(name);

    public IShape Shape(string name) => this.Shapes.GetByName<IShape>(name);

    public void SaveAsPng(Stream stream)
    {
        var imageInfo = new SKImageInfo((int)this.slideSize.Width(), (int)this.slideSize.Height());
        var surface = SKSurface.Create(imageInfo);
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.White); // TODO: #344 get real

        foreach (var autoShape in this.Shapes.OfType<AutoShape>())
        {
            autoShape.Draw(canvas);
        }

        var image = surface.Snapshot();
        var bitmap = SKBitmap.FromImage(image);
        var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
        data.SaveTo(stream);
    }

    public IList<ITextBox> TextFrames()
    {
        var returnList = new List<ITextBox>();

        var frames = this.Shapes
            .Where(x => x.IsTextHolder)
            .Select(t => t.TextBox)
            .ToList();
        returnList.AddRange(frames);

        // if this slide contains a table, the cells from that table will have to be added as well, since they inherit from ITextBoxContainer but are not direct descendants of the slide
        var tablesOnSlide = this.Shapes.OfType<ITable>().ToList();
        if (tablesOnSlide.Any())
        {
            returnList.AddRange(tablesOnSlide.SelectMany(table =>
                table.Rows.SelectMany(row => row.Cells).Select(cell => cell.TextFrame)));
        }

        // if there are groups on that slide, they need to be added as well since those are not direct descendants of the slide either
        var groupsOnSlide = this.Shapes.OfType<IGroupShape>().ToList();
        if (groupsOnSlide.Any())
        {
            foreach (var group in groupsOnSlide)
            {
                this.AddAllTextboxesInGroupToList(group, returnList);
            }
        }

        return returnList;
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
            foreach(var line in lines)
            {
                paragraphs.Add();
                paragraphs[paragraphs.Count - 1].Text = line;
            }
        }
    }

    internal PresentationDocument SDKPresentationDocument() => (PresentationDocument)this.SDKSlidePart.OpenXmlPackage;

    /// <summary>
    ///     Iterates group recursively and add all text boxes in the list.
    /// </summary>
    private void AddAllTextboxesInGroupToList(IGroupShape group, List<ITextBox> textBoxes)
    {
        foreach (var shape in group.Shapes)
        {
            switch (shape.ShapeType)
            {
                case ShapeType.Group:
                    this.AddAllTextboxesInGroupToList((IGroupShape)shape, textBoxes);
                    break;
                case ShapeType.AutoShape:
                    if (shape.IsTextHolder)
                    {
                        textBoxes.Add(shape.TextBox);
                    }

                    break;
            }
        }
    }

    private ITextBox? GetNotes()
    {
        var notes = this.SDKSlidePart.NotesSlidePart;

        if (notes is null)
        {
            return null;
        }

        var shapes = new ShapeCollection.Shapes(notes);
        var notesPlaceholder = shapes
            .FirstOrDefault(x => 
                x.IsPlaceholder && 
                x.IsTextHolder && 
                x.PlaceholderType == PlaceholderType.Text);
        return notesPlaceholder?.TextBox;
    }

    private void AddNotesSlide(IEnumerable<string> lines)
    {
        // Build up the children of the text body element
        var textBodyChildren = new List<OpenXmlElement>() {
            new BodyProperties(),
            new ListStyle()
        };

        // Add in the text lines
        textBodyChildren.AddRange(
            lines
                .Select(line => new A.Paragraph(
                    new A.ParagraphProperties(),
                    new A.Run(
                        new A.RunProperties(),
                        new A.Text(line)),
                    new A.EndParagraphRunProperties())));

        // Always add at least one paragraph, even if empty
        if (!lines.Any())
        {
            textBodyChildren.Add(
                new A.Paragraph(
                    new A.EndParagraphRunProperties()));
        }

        // https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-notes-slides
        var rid = this.SDKSlidePart.NextRelationshipId();
        NotesSlidePart notesSlidePart1 = this.SDKSlidePart.AddNewPart<NotesSlidePart>(rid);
        NotesSlide notesSlide = new NotesSlide(
            new CommonSlideData(
                new ShapeTree(
                    new P.NonVisualGroupShapeProperties(
                        new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = string.Empty },
                        new P.NonVisualGroupShapeDrawingProperties(),
                        new ApplicationNonVisualDrawingProperties()),
                    new GroupShapeProperties(new TransformGroup()),
                    new P.Shape(
                        new P.NonVisualShapeProperties(
                            new P.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Notes Placeholder 2" },
                            new P.NonVisualShapeDrawingProperties(new ShapeLocks() { NoGrouping = true }),
                            new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Body })),
                        new P.ShapeProperties(),
                        new P.TextBody(
                            textBodyChildren)))),           
            new ColorMapOverride(new MasterColorMapping()));
        notesSlidePart1.NotesSlide = notesSlide;
    }
    
    private int ParseNumber()
    {
        var sdkPresentationDocument = (PresentationDocument)this.SDKSlidePart.OpenXmlPackage;
        var presentationPart = sdkPresentationDocument.PresentationPart!;
        var currentSlidePartId = presentationPart.GetIdOfPart(this.SDKSlidePart);
        var slideIdList =
            presentationPart.Presentation.SlideIdList!.ChildElements.OfType<SlideId>().ToList();
        for (int i = 0; i < slideIdList.Count; i++)
        {
            if (slideIdList[i].RelationshipId == currentSlidePartId)
            {
                return i + 1;
            }
        }

        throw new SCException("An error occurred while parsing slide number.");
    }

    private void UpdateNumber(int newSlideNumber)
    {
        if (this.Number == newSlideNumber)
        {
            return;
        }

        var currentIndex = this.Number - 1;
        var destIndex = newSlideNumber - 1;
        var sdkPresentationDocument = (PresentationDocument)this.SDKSlidePart.OpenXmlPackage;
        if (destIndex < 0 || currentIndex >= sdkPresentationDocument.PresentationPart!.SlideParts.Count() ||
            destIndex == currentIndex)
        {
            throw new ArgumentOutOfRangeException(nameof(destIndex));
        }

        var presentationPart = sdkPresentationDocument.PresentationPart!;

        var presentation = presentationPart.Presentation;
        var slideIdList = presentation.SlideIdList!;

        // Get the slide ID of the source slide.
        var sourceSlide = (SlideId)slideIdList.ChildElements[currentIndex];

        SlideId? targetSlide;

        // Identify the position of the target slide after which to move the source slide
        if (destIndex == 0)
        {
            targetSlide = null;
        }
        else if (currentIndex < destIndex)
        {
            targetSlide = (SlideId)slideIdList.ChildElements[destIndex];
        }
        else
        {
            targetSlide = (SlideId)slideIdList.ChildElements[destIndex - 1];
        }

        // Remove the source slide from its current position.
        sourceSlide.Remove();

        // Insert the source slide at its new position after the target slide.
        slideIdList.InsertAfter(sourceSlide, targetSlide);

        // Save the modified presentation.
        presentation.Save();
    }

    private string? GetCustomData()
    {
        if (this.sdkCustomXmlPart.Value == null)
        {
            return null;
        }

        var customXmlPartStream = this.sdkCustomXmlPart.Value.GetStream();
        using var customXmlStreamReader = new StreamReader(customXmlPartStream);
        var raw = customXmlStreamReader.ReadToEnd();
#if NET7_0
        return raw[Constants.CustomDataElementName.Length..];
#else
        return raw.Substring(Constants.CustomDataElementName.Length);
#endif
    }

    private void SetCustomData(string? value)
    {
        Stream customXmlPartStream;
        if (this.sdkCustomXmlPart.Value == null)
        {
            CustomXmlPart newSlideCustomXmlPart = this.SDKSlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPartStream = newSlideCustomXmlPart.GetStream();
            this.sdkCustomXmlPart = new Lazy<CustomXmlPart?>(() => newSlideCustomXmlPart);
        }
        else
        {
            customXmlPartStream = this.sdkCustomXmlPart.Value.GetStream();
        }

        using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
        customXmlStreamReader.Write($"{Constants.CustomDataElementName}{value}");
    }

    private CustomXmlPart? GetSldCustomXmlPart()
    {
        foreach (CustomXmlPart customXmlPart in this.SDKSlidePart.CustomXmlParts)
        {
            using var customXmlPartStream = new StreamReader(customXmlPart.GetStream());
            string customXmlPartText = customXmlPartStream.ReadToEnd();
            if (customXmlPartText.StartsWith(
                    Constants.CustomDataElementName,
                    StringComparison.CurrentCulture))
            {
                return customXmlPart;
            }
        }

        return null;
    }
}