using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler.Collections;

internal class SlideMasterCollection : ISlideMasterCollection
{
    private readonly List<ISlideMaster> slideMasters;

    private SlideMasterCollection(SCPresentation presentation, List<ISlideMaster> slideMasters)
    {
        this.Presentation = presentation;
        this.slideMasters = slideMasters;
    }

    public int Count => this.slideMasters.Count;

    internal SCPresentation Presentation { get; }

    public ISlideMaster this[int index] => this.slideMasters[index];

    public IEnumerator<ISlideMaster> GetEnumerator()
    {
        return this.slideMasters.GetEnumerator();
    }

    internal static SlideMasterCollection Create(SCPresentation presentation)
    {
        IEnumerable<SlideMasterPart> slideMasterParts = presentation.SDKPresentationInternal.PresentationPart!.SlideMasterParts;
        var slideMasters = new List<ISlideMaster>(slideMasterParts.Count());
        foreach (SlideMasterPart slideMasterPart in slideMasterParts)
        {
            slideMasters.Add(new SCSlideMaster(presentation, slideMasterPart.SlideMaster));
        }

        return new SlideMasterCollection(presentation, slideMasters);
    }

    internal SCSlideLayout GetSlideLayoutBySlide(SCSlide slide)
    {
        SlideLayoutPart inputSlideLayoutPart = slide.SDKSlidePart.SlideLayoutPart!;
        IEnumerable<SCSlideLayout> allLayouts = this.slideMasters.SelectMany(sm => sm.SlideLayouts).OfType<SCSlideLayout>();

        return allLayouts.First(sl => sl.SlideLayoutPart.Uri == inputSlideLayoutPart.Uri);
    }
}