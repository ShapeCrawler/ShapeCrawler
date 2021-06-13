using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.SlideMasters;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a collections of Slide Masters.
    /// </summary>
    public interface ISlideMasterCollection
    {
        /// <summary>
        ///     Gets the number of series items in the collection.
        /// </summary>
        int Count { get; }

        /// <summary>
        ///     Gets the element at the specified index.
        /// </summary>
        ISlideMaster this[int index] { get; }

        /// <summary>
        ///     Gets the generic enumerator that iterates through the collection.
        /// </summary>
        IEnumerator<ISlideMaster> GetEnumerator();
    }

    internal class SlideMasterCollection : ISlideMasterCollection // TODO: add interface
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
            IEnumerable<SlideMasterPart> slideMasterParts = presentation.PresentationDocument.PresentationPart.SlideMasterParts;
            var slideMasters = new List<ISlideMaster>(slideMasterParts.Count());
            foreach (SlideMasterPart slideMasterPart in slideMasterParts)
            {
                slideMasters.Add(new SCSlideMaster(presentation, slideMasterPart.SlideMaster));
            }

            return new SlideMasterCollection(presentation, slideMasters);
        }

        internal SCSlideLayout GetSlideLayoutBySlide(SCSlide slide)
        {
            SlideLayoutPart inputSlideLayoutPart = slide.SlidePart.SlideLayoutPart;

            ISlideLayout slideLayout = this.slideMasters.SelectMany(sm => sm.SlideLayouts)
                .First(sl => ((SCSlideLayout) sl).SlideLayoutPart == inputSlideLayoutPart);

            return (SCSlideLayout)slideLayout;
        }
    }
}