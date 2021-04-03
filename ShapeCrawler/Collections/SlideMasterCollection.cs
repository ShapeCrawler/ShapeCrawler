using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Collections
{
    public class SlideMasterCollection : LibraryCollection<SCSlideMaster> //TODO: add interface
    {
        private SlideMasterCollection(SCPresentation presentation, List<SCSlideMaster> slideMasters)
        {
            Presentation = presentation;
            CollectionItems = slideMasters;
        }

        internal SCPresentation Presentation { get; }

        internal static SlideMasterCollection Create(SCPresentation presentation)
        {
            IEnumerable<SlideMasterPart> slideMasterParts = presentation.PresentationPart.SlideMasterParts;
            var slideMasters = new List<SCSlideMaster>(slideMasterParts.Count());
            foreach (SlideMasterPart slideMasterPart in slideMasterParts)
            {
                slideMasters.Add(new SCSlideMaster(presentation, slideMasterPart.SlideMaster));
            }

            return new SlideMasterCollection(presentation, slideMasters);
        }

        internal SCSlideLayout GetSlideLayoutBySlide(SCSlide slide)
        {
            SlideLayoutPart inputSlideLayoutPart = slide.SlidePart.SlideLayoutPart;

            return CollectionItems.SelectMany(sm => sm.SlideLayouts)
                .First(sl => sl.SlideLayoutPart == inputSlideLayoutPart);
        }

        internal SCSlideMaster GetSlideMasterByLayout(SCSlideLayout slideLayout)
        {
            return CollectionItems.First(sldMaster =>
                sldMaster.SlideLayouts.Any(sl => sl.SlideLayoutPart == slideLayout.SlideLayoutPart));
        }
    }
}