using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a slide collection.
    /// </summary>
    public class SlideCollection : ISlideCollection
    {
        private readonly SCPresentation _presentation;
        private readonly PresentationPart _presentationPart;
        private readonly ResettableLazy<List<SlideSc>> _slides;

        internal SlideCollection(SCPresentation presentation)
        {
            _presentationPart = presentation.PresentationPart;
            _presentation = presentation;

            _slides = new ResettableLazy<List<SlideSc>>(GetSlides);
        }

        public IEnumerator<SlideSc> GetEnumerator()
        {
            return _slides.Value.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public SlideSc this[int index] => _slides.Value[index];

        public int Count => _slides.Value.Count;

        /// <summary>
        ///     Removes the specified slide.
        /// </summary>
        public void Remove(SlideSc removingSlide)
        {
            P.Presentation presentation = _presentationPart.Presentation;

            // Get the list of slide identifiers in the presentation
            P.SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide identifier of the specified slide
            P.SlideId slideId = (P.SlideId) slideIdList.ChildElements[removingSlide.Number - 1];

            // Gets the relationship identifier of the slide
            string slideRelId = slideId.RelationshipId;

            // Remove the slide from the slide list
            slideIdList.RemoveChild(slideId);

            // Remove references to the slide from all custom shows
            if (presentation.CustomShowList != null)
            {
                // Iterate through the list of custom shows
                foreach (var customShow in presentation.CustomShowList.Elements<P.CustomShow>())
                {
                    if (customShow.SlideList == null)
                    {
                        continue;
                    }

                    // declares a link list of slide list entries
                    var slideListEntries = new LinkedList<P.SlideListEntry>();
                    foreach (P.SlideListEntry slideListEntry in customShow.SlideList.Elements())
                    {
                        // finds the slide reference to remove from the custom show
                        if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                        {
                            slideListEntries.AddLast(slideListEntry);
                        }
                    }

                    // Removes all references to the slide from the custom show
                    foreach (P.SlideListEntry slideListEntry in slideListEntries)
                    {
                        customShow.SlideList.RemoveChild(slideListEntry);
                    }
                }
            }

            // Gets the slide part for the specified slide
            SlidePart slidePart = _presentationPart.GetPartById(slideRelId) as SlidePart;

            _presentationPart.DeletePart(slidePart);
            _presentationPart.Presentation.Save();
            _slides.Reset();
        }

        private List<SlideSc> GetSlides()
        {
            int slidesCount = _presentationPart.SlideParts.Count();
            var slides = new List<SlideSc>(slidesCount);
            for (var sldIndex = 0; sldIndex < slidesCount; sldIndex++)
            {
                SlidePart slidePart = _presentationPart.GetSlidePartByIndex(sldIndex);
                var slideNumber = new SlideNumber(sldIndex + 1);
                var newSlide = new SlideSc(_presentation, slidePart, slideNumber);
                slides.Add(newSlide);
            }

            return slides;
        }
    }
}