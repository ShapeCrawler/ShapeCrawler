using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using AngleSharp;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using SkiaSharp;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler;

internal sealed class SCSlide : ISlide
{
    private readonly ResetableLazy<SCSlideShapes> shapes;
    private readonly Lazy<SCImage?> backgroundImage;
    private readonly Func<int> totalSlideCount;
    private Lazy<CustomXmlPart?> customXmlPart;
    private readonly int slideWidth;
    private readonly int slideHeight;
    private readonly SlidePart sdkSlidePart;

    internal SCSlide( 
        SlidePart slidePart, 
        SlideId slideId, 
        Func<int> totalSlideCount,
        SCSlideLayout slideLayout,
        int slideWidth, 
        int slideHeight,
        PresentationDocument sdkPresentationDocument)
    {
        this.sdkSlidePart = slidePart;
        this.slideWidth = slideWidth;
        this.slideHeight = slideHeight;
        this.shapes = new ResetableLazy<SCSlideShapes>(() => new SCSlideShapes(this.sdkSlidePart, this, slidePart, imageParts, sdkPresentationDocument));
        this.backgroundImage = new Lazy<SCImage?>(() => SCImage.ForBackground(this, this.imageParts));
        this.customXmlPart = new Lazy<CustomXmlPart?>(this.GetSldCustomXmlPart);
        this.SlideId = slideId;
        this.totalSlideCount = totalSlideCount;
        this.SlideLayout = slideLayout;
        this.SDKPresentationDocument = sdkPresentationDocument;
    }

    public ISlideLayout SlideLayout { get; }

    public ISlideShapeCollection Shapes => this.shapes.Value;

    public int Number
    {
        get => this.GetNumber();
        set => this.UpdateNumber(value);
    }

    public IImage? Background => this.backgroundImage.Value;

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value);
    }

    public bool Hidden => this.sdkSlidePart.Slide.Show is not null && this.sdkSlidePart.Slide.Show.Value == false;
    
    public PresentationDocument SDKPresentationDocument { get; }
    
    internal SCSlideLayout SlideLayoutInternal => (SCSlideLayout)this.SlideLayout;

    internal TypedOpenXmlPart TypedOpenXmlPart => this.sdkSlidePart;

    internal SlideId SlideId { get; }

    internal SlidePart SDKSlidePart()
    {
        return this.sdkSlidePart;
    }

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

    public async Task<string> ToHtml()
    {
        var browsingContext = BrowsingContext.New(Configuration.Default.WithDefaultLoader().WithCss());
        var document = await browsingContext.OpenNewAsync().ConfigureAwait(false);
        var body = document.Body!;
        
        foreach (var shape in this.Shapes.OfType<SCShape>())
        {
            body.AppendChild(shape.ToHtmlElement());
        }

        return document.DocumentElement.OuterHtml;
    }

    public void SaveAsPng(Stream stream)
    {
        var imageInfo = new SKImageInfo(this.slideWidth, this.slideHeight);
        var surface = SKSurface.Create(imageInfo);
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.White); // TODO: #344 get real
        
        foreach (var autoShape in this.Shapes.OfType<SCSlideAutoShape>())
        {
            autoShape.Draw(canvas);
        }
        
        var image = surface.Snapshot();
        var bitmap = SKBitmap.FromImage(image);
        var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
        data.SaveTo(stream);
    }

    public IList<ITextFrame> GetAllTextFrames()
    {
        var returnList = new List<ITextFrame>();

        var frames = this.Shapes.OfType<ITextFrameContainer>()
            .Where(t => t.TextFrame != null)
            .Select(t => t.TextFrame!)
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
                case SCShapeType.Group:
                    this.AddAllTextboxesInGroupToList((IGroupShape)shape, textBoxes);
                    break;
                case SCShapeType.AutoShape:
                    if (shape is ITextFrameContainer)
                    {
                        textBoxes.Add(((ITextFrameContainer)shape).TextFrame!);
                    }

                    break;
            }
        }
    }

    private int GetNumber()
    {
        var presentationPart = this.PresCore.SDKPresentation.PresentationPart!;
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
        
        if (destIndex < 0 || currentIndex >= totalSlideCount.Invoke() || destIndex == currentIndex)
        {
            throw new ArgumentOutOfRangeException(nameof(destIndex));
        }

        var presentationPart = this.Presentation.SDKPresentationDocument.PresentationPart!;

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
        if (this.customXmlPart.Value == null)
        {
            return null;
        }

        var customXmlPartStream = this.customXmlPart.Value.GetStream();
        using var customXmlStreamReader = new StreamReader(customXmlPartStream);
        var raw = customXmlStreamReader.ReadToEnd();
#if NET7_0
        return raw[SCConstants.CustomDataElementName.Length..];
#else
        return raw.Substring(SCConstants.CustomDataElementName.Length);
#endif
    }

    private void SetCustomData(string? value)
    {
        Stream customXmlPartStream;
        if (this.customXmlPart.Value == null)
        {
            CustomXmlPart newSlideCustomXmlPart = this.sdkSlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
            customXmlPartStream = newSlideCustomXmlPart.GetStream();
            this.customXmlPart = new Lazy<CustomXmlPart?>(() => newSlideCustomXmlPart);
        }
        else
        {
            customXmlPartStream = this.customXmlPart.Value.GetStream();
        }

        using var customXmlStreamReader = new StreamWriter(customXmlPartStream);
        customXmlStreamReader.Write($"{SCConstants.CustomDataElementName}{value}");
    }

    private CustomXmlPart? GetSldCustomXmlPart()
    {
        foreach (CustomXmlPart customXmlPart in this.sdkSlidePart.CustomXmlParts)
        {
            using var customXmlPartStream = new StreamReader(customXmlPart.GetStream());
            string customXmlPartText = customXmlPartStream.ReadToEnd();
            if (customXmlPartText.StartsWith(
                SCConstants.CustomDataElementName,
                StringComparison.CurrentCulture))
            {
                return customXmlPart;
            }
        }

        return null;
    }
}