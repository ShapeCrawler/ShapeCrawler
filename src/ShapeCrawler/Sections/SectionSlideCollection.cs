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
    private readonly PresentationDocument sdkPresDocument;
    private readonly IEnumerable<SectionSlideIdListEntry> p14SectionSlideIdListEntryList;

    internal SectionSlideCollection(
        PresentationDocument sdkPresDocument,
        IEnumerable<SectionSlideIdListEntry> p14SectionSlideIdListEntryList)
    {
        this.sdkPresDocument = sdkPresDocument;
        this.p14SectionSlideIdListEntryList = p14SectionSlideIdListEntryList;
    }

    public int Count => this.ReadOnlySlides().Count;

    public ISlide this[int index] => this.ReadOnlySlides()[index];

    public IEnumerator<ISlide> GetEnumerator() => this.ReadOnlySlides().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private ReadOnlySlides ReadOnlySlides()
    {
        var sdkSlideParts = new List<SlidePart>();
        var idToRId = this.sdkPresDocument.PresentationPart!.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>()
            .ToDictionary(x => x.Id!, x => x.RelationshipId);
        foreach (var p14SectionSlideIdListEntry in this.p14SectionSlideIdListEntryList)
        {
            var rId = idToRId[p14SectionSlideIdListEntry.Id!]!.Value!;
            var slidePart = (SlidePart)this.sdkPresDocument.PresentationPart!.GetPartById(rId);
            sdkSlideParts.Add(slidePart);
        }

        return new ReadOnlySlides(sdkSlideParts);
    }
}