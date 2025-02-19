using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Presentations;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed record ReadOnlySlides : IReadOnlyList<ISlide>
{
    private readonly IEnumerable<SlidePart> slideParts;

    private readonly MediaCollection mediaCollection = new();

    internal ReadOnlySlides(IEnumerable<SlidePart> slideParts)
    {
        this.slideParts = slideParts;
        this.BuildMediaCollection();
    }

    public int Count => this.SlideList().Count;

    public ISlide this[int index] => this.SlideList()[index];

    public IEnumerator<ISlide> GetEnumerator() => this.SlideList().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<Slide> SlideList()
    {
        if (!this.slideParts.Any())
        {
            return [];
        }

        var presDocument = (PresentationDocument)this.slideParts.First().OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var slidesCount = this.slideParts.Count();
        var slides = new List<Slide>(slidesCount);
        var pSlideIdList = presPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>().ToList();
        for (var slideIndex = 0; slideIndex < slidesCount; slideIndex++)
        {
            var pSlideId = pSlideIdList[slideIndex];
            var sdkSlidePart = (SlidePart)presPart.GetPartById(pSlideId.RelationshipId!);
            var layout = new SlideLayout(sdkSlidePart.SlideLayoutPart!);
            var newSlide = new Slide(sdkSlidePart, layout, this.mediaCollection);
            slides.Add(newSlide);
        }

        return slides;
    }

    private void BuildMediaCollection()
    {
        var imageParts = this.slideParts.SelectMany(x => x.ImageParts);
        foreach (var imagePart in imageParts)
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