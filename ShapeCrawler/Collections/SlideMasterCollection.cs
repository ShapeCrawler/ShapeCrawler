using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Collections
{
    public class SlideMasterCollection : LibraryCollection<SlideMasterSc> //TODO: add interface
    {
        internal PresentationSc Presentation { get; }

        private SlideMasterCollection(PresentationSc presentation, List<SlideMasterSc> slideMasters)
        {
            Presentation = presentation;
            CollectionItems = slideMasters;
        }

        internal static SlideMasterCollection Create(PresentationSc presentation, IEnumerable<SlideMasterPart> slideMasterParts)
        {
            var slideMasters = new List<SlideMasterSc>(slideMasterParts.Count());
            foreach (SlideMasterPart slideMasterPart in slideMasterParts)
            {
                slideMasters.Add(new SlideMasterSc(presentation, slideMasterPart.SlideMaster));
            }

            return new SlideMasterCollection(presentation, slideMasters);
        }

        internal SlideLayoutSc GetSlideLayoutBySlide(SlideSc slide)
        {
            return new SlideLayoutSc(slide.SlidePart.SlideLayoutPart, slide.Presentation);
        }

        internal SlideMasterSc GetSlideMasterByLayout(SlideLayoutSc slideLayoutSc)
        {
            return CollectionItems.First(sldMaster =>
                sldMaster.SlideLayouts.Any(sldLayout => sldLayout.SlideLayoutPart == slideLayoutSc.SlideLayoutPart));
        }
    }
}