using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Presentations;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideCollection : IReadOnlyList<ISlide>
{
    private readonly IEnumerable<SlidePart> slideParts;
    private readonly MediaCollection mediaCollection = new();

    internal SlideCollection(IEnumerable<SlidePart> slideParts)
    {
        this.slideParts = slideParts;
        this.GetMediaCollection();
    }

    public int Count => this.GetSlides().Count;

    public ISlide this[int index] => this.GetSlides()[index];

    public IEnumerator<ISlide> GetEnumerator() => this.GetSlides().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<Slide> GetSlides()
    {
        if (!this.slideParts.Any())
        {
            return [];
        }

        var presDocument = (PresentationDocument)this.slideParts.First().OpenXmlPackage;
        var presPart = presDocument.PresentationPart!;
        var pSlideIdList = presPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>().ToList();
        var slidesCount = pSlideIdList.Count;
        var slides = new List<Slide>(slidesCount);
        foreach (var pSlideId in pSlideIdList)
        {
            var slidePart = (SlidePart)presPart.GetPartById(pSlideId.RelationshipId!);
            var newSlide = new Slide(
                slidePart,
                new SlideLayout(slidePart.SlideLayoutPart!),
                this.mediaCollection);
            slides.Add(newSlide);
        }

        return slides;
    }

    private void GetMediaCollection()
    {
        var imageParts = this.slideParts.SelectMany(slidePart => slidePart.ImageParts);
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