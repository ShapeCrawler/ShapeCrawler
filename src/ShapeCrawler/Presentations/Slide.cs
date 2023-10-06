using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AngleSharp;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Drawing;
using ShapeCrawler.Exceptions;
using ShapeCrawler.ShapeCollection;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideShape;
using SkiaSharp;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler;

internal sealed class Slide : ISlide
{
    private readonly Lazy<SlideBgImage> backgroundImage;
    private Lazy<CustomXmlPart?> sdkCustomXmlPart;
    private readonly SlidePart sdkSlidePart;
    private readonly SlideSize slideSize;

    internal Slide(
        SlidePart sdkSlidePart,
        ISlideLayout slideLayout,
        SlideSize slideSize)
    {
        this.sdkSlidePart = sdkSlidePart;
        this.slideSize = slideSize;
        this.backgroundImage = new Lazy<SlideBgImage>(() =>
            new SlideBgImage(sdkSlidePart));
        this.sdkCustomXmlPart = new Lazy<CustomXmlPart?>(this.GetSldCustomXmlPart);
        this.SlideLayout = slideLayout;
        this.Shapes = new SlideShapes(this.sdkSlidePart, new ShapeCollection.Shapes(sdkSlidePart));
    }

    public ISlideLayout SlideLayout { get; }

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

    public bool Hidden() => this.sdkSlidePart.Slide.Show is not null && this.sdkSlidePart.Slide.Show.Value == false;

    public void Hide()
    {
        if (this.sdkSlidePart.Slide.Show is null)
        {
            var showAttribute = new OpenXmlAttribute("show", string.Empty, "0");
            this.sdkSlidePart.Slide.SetAttribute(showAttribute);
        }
        else
        {
            this.sdkSlidePart.Slide.Show = false;
        }
    }

    public IShape ShapeWithName(string shape) => this.Shapes.GetByName<IShape>(shape);

    public ITable TableWithName(string table) => this.Shapes.GetByName<ITable>(table);

    public async Task<string> ToHtml()
    {
        var browsingContext = BrowsingContext.New(Configuration.Default.WithDefaultLoader().WithCss());
        var document = await browsingContext.OpenNewAsync().ConfigureAwait(false);
        var body = document.Body!;

        foreach (var shape in this.Shapes.OfType<AutoShape>())
        {
            body.AppendChild(shape.ToHtmlElement());
        }

        return document.DocumentElement.OuterHtml;
    }

    public void SaveAsPng(Stream stream)
    {
        var imageInfo = new SKImageInfo(this.slideSize.Width(), this.slideSize.Height());
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

    public IList<ITextFrame> TextFrames()
    {
        var returnList = new List<ITextFrame>();

        var frames = this.Shapes
            .Where(x => x.IsTextHolder)
            .Select(t => t.TextFrame)
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

    /// <summary>
    ///     Iterates group recursively and add all text boxes in the list.
    /// </summary>
    private void AddAllTextboxesInGroupToList(IGroupShape group, List<ITextFrame> textBoxes)
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
                        textBoxes.Add(shape.TextFrame);
                    }

                    break;
            }
        }
    }

    private int ParseNumber()
    {
        var sdkPresentationDocument = (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
        var presentationPart = sdkPresentationDocument.PresentationPart!;
        var currentSlidePartId = presentationPart.GetIdOfPart(this.sdkSlidePart);
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
        var sdkPresentationDocument = (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
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
        return raw.Substring(SCConstants.CustomDataElementName.Length);
#endif
    }

    private void SetCustomData(string? value)
    {
        Stream customXmlPartStream;
        if (this.sdkCustomXmlPart.Value == null)
        {
            CustomXmlPart newSlideCustomXmlPart = this.sdkSlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
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
        foreach (CustomXmlPart customXmlPart in this.sdkSlidePart.CustomXmlParts)
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

    internal PresentationDocument SDKPresentationDocument()
    {
        return (PresentationDocument)this.sdkSlidePart.OpenXmlPackage;
    }
}