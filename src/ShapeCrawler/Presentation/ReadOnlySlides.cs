using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler;

internal sealed record ReadOnlySlides : IReadOnlyList<ISlide>
{
    private readonly IEnumerable<SlidePart> sdkSlideParts;

    internal ReadOnlySlides(IEnumerable<SlidePart> sdkSlideParts)
    {
        this.sdkSlideParts = sdkSlideParts;
    }

    public int Count => this.SlideList().Count;

    public ISlide this[int index] => this.SlideList()[index];

    public IEnumerator<ISlide> GetEnumerator()=>this.SlideList().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator()=> this.GetEnumerator();
    
    private List<Slide> SlideList()
    {
        if (!this.sdkSlideParts.Any())
        {
            return new List<Slide>(0);
        }
        
        var sdkPresDocument = (PresentationDocument)this.sdkSlideParts.First().OpenXmlPackage;
        var sdkPresPart = sdkPresDocument.PresentationPart!;
        int slidesCount = this.sdkSlideParts.Count();
        var slides = new List<Slide>(slidesCount);
        var pSlideIdList = sdkPresPart.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>().ToList();
        for (var slideIndex = 0; slideIndex < slidesCount; slideIndex++)
        {
            var pSlideId = pSlideIdList[slideIndex];
            var sdkSlidePart = (SlidePart)sdkPresPart.GetPartById(pSlideId.RelationshipId!);
            var layout = new SlideLayout(sdkSlidePart.SlideLayoutPart!);
            var slideSize = new SlideSize(sdkPresDocument.PresentationPart!.Presentation.SlideSize!);
            var newSlide = new Slide(sdkSlidePart, pSlideId, layout, slideSize);
            slides.Add(newSlide);
        }

        return slides;
    }
}