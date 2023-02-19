using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Constants;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Shared;
using ShapeCrawler.SlideMasters;
using ShapeCrawler.Texts;
using SkiaSharp;

// ReSharper disable CheckNamespace
// ReSharper disable PossibleMultipleEnumeration
namespace ShapeCrawler;

internal sealed class SCSlide : SlideStructure, ISlide
{
    private readonly ResettableLazy<ShapeCollection> shapes;
    private readonly Lazy<SCImage?> backgroundImage;
    private Lazy<CustomXmlPart?> customXmlPart;

    internal SCSlide(SCPresentation pres, SlidePart slidePart, SlideId slideId)
    : base(pres)
    {
        this.Presentation = pres;
        this.SDKSlidePart = slidePart;
        this.shapes = new ResettableLazy<ShapeCollection>(() => new ShapeCollection(this.SDKSlidePart, this));
        this.backgroundImage = new Lazy<SCImage?>(() => SCImage.ForBackground(this));
        this.customXmlPart = new Lazy<CustomXmlPart?>(this.GetSldCustomXmlPart);
        this.SlideId = slideId;
    }

    public ISlideLayout SlideLayout =>
        ((SlideMasterCollection)this.PresentationInternal.SlideMasters).GetSlideLayoutBySlide(this);

    public IShapeCollection Shapes => this.shapes.Value;

    public override int Number
    {
        get => this.GetNumber();
        set => this.SetNumber(value);
    }

    public IImage? Background => this.backgroundImage.Value;

    public string? CustomData
    {
        get => this.GetCustomData();
        set => this.SetCustomData(value);
    }

    public bool Hidden => this.SDKSlidePart.Slide.Show is not null && this.SDKSlidePart.Slide.Show.Value == false;

    public SlidePart SDKSlidePart { get; }

    internal SCSlideLayout SlideLayoutInternal => (SCSlideLayout)this.SlideLayout;

    internal override TypedOpenXmlPart TypedOpenXmlPart => this.SDKSlidePart;

    internal SlideId SlideId { get; }

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

    public void SaveAsPng(Stream stream)
    {
        var imageInfo = new SKImageInfo(this.PresentationInternal.SlideWidth, this.PresentationInternal.SlideHeight);
        var surface = SKSurface.Create(imageInfo);
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.White); // TODO: #344 get real
        
        foreach (var autoShape in this.Shapes.OfType<SCAutoShape>())
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
        List<ITextFrame> returnList = new List<ITextFrame>();

        // this will add all textboxes from shapes on that slide that directly inherit ITextBoxContainer
        returnList.AddRange(this.Shapes.OfType<ITextFrameContainer>()
            .Where(t => t.TextFrame != null)
            .Select(t => t.TextFrame)
            .ToList());

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
                case SCShapeType.GroupShape:
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
        var presentationPart = this.PresentationInternal.SDKPresentationInternal.PresentationPart!;
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

    private void SetNumber(int newSlideNumber)
    {
        int from = this.Number - 1;
        int to = newSlideNumber - 1;

        if (to < 0 || from >= this.PresentationInternal.Slides.Count || to == from)
        {
            throw new ArgumentOutOfRangeException(nameof(to));
        }

        var presentationPart = this.PresentationInternal.SDKPresentationInternal.PresentationPart!;

        var presentation = presentationPart.Presentation;
        var slideIdList = presentation.SlideIdList!;

        // Get the slide ID of the source slide.
        var sourceSlide = (SlideId)slideIdList.ChildElements[from];

        SlideId? targetSlide;

        // Identify the position of the target slide after which to move the source slide
        if (to == 0)
        {
            targetSlide = null;
        }
        else if (from < to)
        {
            targetSlide = (SlideId)slideIdList.ChildElements[to];
        }
        else
        {
            targetSlide = (SlideId)slideIdList.ChildElements[to - 1];
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
            CustomXmlPart newSlideCustomXmlPart = this.SDKSlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
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
        foreach (CustomXmlPart customXmlPart in this.SDKSlidePart.CustomXmlParts)
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