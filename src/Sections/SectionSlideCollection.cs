using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Slides;
using P = DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler.Sections;

internal sealed class SectionSlideCollection(P14.Section p14Section) : IReadOnlyList<IUserSlide>
{
    public int Count => this.GetSlides().Count;

    public IUserSlide this[int index] => this.GetSlides()[index];

    public IEnumerator<IUserSlide> GetEnumerator() => this.GetSlides().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.GetEnumerator();

    private List<IUserSlide> GetSlides()
    {
        var sectionIds = p14Section.Descendants<P14.SectionSlideIdListEntry>()
            .Select(entry => entry.Id!.Value).ToHashSet();
        var presDocument = new SCOpenXmlElement(p14Section).ParentPresentationDocument;
        var presentationPart = presDocument.PresentationPart!;
        var slideIds = presentationPart.Presentation!.SlideIdList!.Elements<P.SlideId>().ToList();
        return [.. new UserSlideCollection(presentationPart.SlideParts)
            .Where(slide => sectionIds.Contains(slideIds[slide.Number - 1].Id!.Value))];
    }
}
