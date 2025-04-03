using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Sections;

internal sealed class SectionSlideCollection(P14.Section p14Section): IReadOnlyList<ISlide>
{
    public int Count => this.GetSlides().Count;

    public ISlide this[int index] => this.GetSlides()[index];

    public IEnumerator<ISlide> GetEnumerator() => this.GetSlides().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private SlideCollection GetSlides()
    {
        var p14SectionSlideIdListEntryList = p14Section.Descendants<P14.SectionSlideIdListEntry>();
        var presDocument = new SCOpenXmlElement(p14Section).PresentationDocument;
        var slideParts = new List<SlidePart>();
        var idToRId = presDocument.PresentationPart!.Presentation.SlideIdList!.ChildElements.OfType<P.SlideId>()
            .ToDictionary(slideId => slideId.Id!, slideId => slideId.RelationshipId);
        foreach (var p14SectionSlideIdListEntry in p14SectionSlideIdListEntryList)
        {
            var rId = idToRId[p14SectionSlideIdListEntry.Id!]!.Value!;
            var slidePart = (SlidePart)presDocument.PresentationPart!.GetPartById(rId);
            slideParts.Add(slidePart);
        }

        return new SlideCollection(slideParts);
    }
}