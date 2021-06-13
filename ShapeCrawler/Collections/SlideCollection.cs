using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    /// <summary>
    ///     Represents a slide collection.
    /// </summary>
    internal class SlideCollection : ISlideCollection // TODO: make internal
    {
        private readonly SCPresentation parentPresentation;
        private PresentationPart presentationPart;
        private readonly ResettableLazy<List<SCSlide>> slides;

        internal SlideCollection(SCPresentation presentation)
        {
            this.presentationPart = presentation.PresentationPart;
            this.parentPresentation = presentation;
            this.slides = new ResettableLazy<List<SCSlide>>(this.GetSlides);
        }

        public int Count => this.slides.Value.Count;

        public ISlide this[int index] => this.slides.Value[index];

        public IEnumerator<ISlide> GetEnumerator()
        {
            return this.slides.Value.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public void Remove(ISlide removingSlide)
        {
            P.Presentation presentation = this.presentationPart.Presentation;

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
            SlidePart slidePart = this.presentationPart.GetPartById(slideRelId) as SlidePart;

            this.presentationPart.DeletePart(slidePart);
            this.presentationPart.Presentation.Save();
            removingSlide.IsRemoved = true;

            this.slides.Reset();
        }

        public void Add(ISlide addingSlide)
        {
            throw new NotImplementedException();
        }

        private List<SCSlide> GetSlides()
        {
            this.presentationPart = this.parentPresentation.presentationDocument.PresentationPart;
            int slidesCount = this.presentationPart.SlideParts.Count();
            var slides = new List<SCSlide>(slidesCount);
            for (var slideIndex = 0; slideIndex < slidesCount; slideIndex++)
            {
                SlidePart slidePart = this.presentationPart.GetSlidePartByIndex(slideIndex);
                int slideNumber = slideIndex + 1;
                var newSlide = new SCSlide(this.parentPresentation, slidePart, slideNumber);
                slides.Add(newSlide);
            }

            return slides;
        }
    }
}