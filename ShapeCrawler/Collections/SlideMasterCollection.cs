using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Collections
{
    public class SlideMasterCollection : LibraryCollection<SlideMasterSc> //TODO: add interface
    {
        private SlideMasterCollection(SCPresentation presentation, List<SlideMasterSc> slideMasters)
        {
            Presentation = presentation;
            CollectionItems = slideMasters;
        }

        internal SCPresentation Presentation { get; }

        internal static SlideMasterCollection Create(SCPresentation presentation)
        {
            IEnumerable<SlideMasterPart> slideMasterParts = presentation.PresentationPart.SlideMasterParts;
            var slideMasters = new List<SlideMasterSc>(slideMasterParts.Count());
            foreach (SlideMasterPart slideMasterPart in slideMasterParts)
            {
                slideMasters.Add(new SlideMasterSc(presentation, slideMasterPart.SlideMaster));
            }

            return new SlideMasterCollection(presentation, slideMasters);
        }

        internal SlideLayoutSc GetSlideLayoutBySlide(SlideSc slide)
        {
            SlideLayoutPart inputSlideLayoutPart = slide.SlidePart.SlideLayoutPart;

            return CollectionItems.SelectMany(sm => sm.SlideLayouts)
                .First(sl => sl.SlideLayoutPart == inputSlideLayoutPart);
        }

        internal SlideMasterSc GetSlideMasterByLayout(SlideLayoutSc slideLayout)
        {
            return CollectionItems.First(sldMaster =>
                sldMaster.SlideLayouts.Any(sl => sl.SlideLayoutPart == slideLayout.SlideLayoutPart));
        }
    }
}