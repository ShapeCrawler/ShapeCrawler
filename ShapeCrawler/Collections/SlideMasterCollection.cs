using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Collections
{
    public class SlideMasterCollection : LibraryCollection<SlideMasterSc>
    {
        private SlideMasterCollection(List<SlideMasterSc> slideMasters)
        {
            CollectionItems = slideMasters;
        }

        public static SlideMasterCollection Create(IEnumerable<SlideMasterPart> slideMasterParts)
        {
            var slideMasters = new List<SlideMasterSc>(slideMasterParts.Count());
            foreach (SlideMasterPart slideMasterPart in slideMasterParts)
            {
                slideMasters.Add(new SlideMasterSc(slideMasterPart.SlideMaster));
            }

            return new SlideMasterCollection(slideMasters);
        }

        public SlideLayoutSc GetSlideLayout(SlideLayoutPart slideLayoutPart)
        {
            throw new System.NotImplementedException();
        }

        public SlideMasterSc GetSlideMaster(SlideLayoutPart slideLayoutPart)
        {
            throw new System.NotImplementedException();
        }
    }
}