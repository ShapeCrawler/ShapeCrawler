using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.ShapeCollection;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Presentations;

internal sealed record ReadOnlySlides : IReadOnlyList<ISlide>
{
    private readonly IEnumerable<SlidePart> sdkSlideParts;

    private readonly MediaCollection mediaCollection = new();

    internal ReadOnlySlides(IEnumerable<SlidePart> sdkSlideParts)
    {
        this.sdkSlideParts = sdkSlideParts;
        this.BuildMediaCollection();
    }

    public int Count => this.SlideList().Count;

    public ISlide this[int index] => this.SlideList()[index];

    public IEnumerator<ISlide> GetEnumerator() => this.SlideList().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<Slide> SlideList()
    {
        if (!this.sdkSlideParts.Any())
        {
            return [];
        }

        var sdkPresDocument = (PresentationDocument)this.sdkSlideParts.First().OpenXmlPackage;
        var sdkPresPart = sdkPresDocument.PresentationPart!;
        var slidesCount = this.sdkSlideParts.Count();
        var slides = new List<Slide>(slidesCount);
        var pSlideIdList = sdkPresPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>().ToList();
        for (var slideIndex = 0; slideIndex < slidesCount; slideIndex++)
        {
            var pSlideId = pSlideIdList[slideIndex];
            var sdkSlidePart = (SlidePart)sdkPresPart.GetPartById(pSlideId.RelationshipId!);
            var layout = new SlideLayout(sdkSlidePart.SlideLayoutPart!);
            var slideSize = new SlideSize(sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
            var newSlide = new Slide(sdkSlidePart, layout, slideSize, this.mediaCollection);
            slides.Add(newSlide);
        }

        return slides;
    }

    private void BuildMediaCollection()
    {
        var imageParts = this.sdkSlideParts.SelectMany(x => x.ImageParts);
        foreach(var imagePart in imageParts)
        {
            using var stream = imagePart.GetStream();
            var hash = new ImageStream(stream).Base64Hash;
            if (!this.mediaCollection.TryGetImagePart(hash, out _))
            {
                this.mediaCollection.SetImagePart(hash, imagePart);
            }
        }
    }
}