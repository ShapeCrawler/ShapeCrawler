using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Sections;

internal sealed class SectionSlideCollection : IReadOnlyList<ISlide>
{
    private readonly PresentationDocument presDocument;
    private readonly IEnumerable<SectionSlideIdListEntry> p14SectionSlideIdListEntryList;

    internal SectionSlideCollection(
        PresentationDocument presDocument,
        IEnumerable<SectionSlideIdListEntry> p14SectionSlideIdListEntryList)
    {
        this.presDocument = presDocument;
        this.p14SectionSlideIdListEntryList = p14SectionSlideIdListEntryList;
    }

    public int Count => this.GetSlides().Count;

    public ISlide this[int index] => this.GetSlides()[index];

    public IEnumerator<ISlide> GetEnumerator() => this.GetSlides().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private ReadOnlySlides GetSlides()
    {
        var slideParts = new List<SlidePart>();
        var idToRId = this.presDocument.PresentationPart!.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>()
            .ToDictionary(slideId => slideId.Id!, slideId => slideId.RelationshipId);
        foreach (var p14SectionSlideIdListEntry in this.p14SectionSlideIdListEntryList)
        {
            var rId = idToRId[p14SectionSlideIdListEntry.Id!]!.Value!;
            var slidePart = (SlidePart)this.presDocument.PresentationPart!.GetPartById(rId);
            slideParts.Add(slidePart);
        }

        return new ReadOnlySlides(slideParts);
    }
}