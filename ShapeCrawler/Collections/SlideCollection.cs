using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Extensions;
using ShapeCrawler.Models;
using ShapeCrawler.Settings;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    /// <summary>
    /// Represents a slide collection.
    /// </summary>
    public class SlideCollection : EditableCollection<SlideSc>
    {
        #region Fields

        private readonly PresentationPart _presentationPart;
        
        // TODO: Consider deleting implementation without this dictionary
        private readonly Dictionary<SlideSc, SlideNumber> _slideToSlideNumber;

        #endregion Fields

        #region Constructors

        private SlideCollection(
            List<SlideSc> slides, 
            PresentationPart presentationPart, 
            Dictionary<SlideSc, SlideNumber> slideToSlideNumber)
        {
            CollectionItems = slides;
            _presentationPart = presentationPart;
            _slideToSlideNumber = slideToSlideNumber;
        }

        #endregion Constructors

        /// <summary>
        /// Removes the specified slide.
        /// </summary>
        /// <param name="row"></param>
        public override void Remove(SlideSc row)
        {
            Check.NotNull(row, nameof(row));

            RemoveFromDom(row.Number);
            _presentationPart.Presentation.Save();
            CollectionItems.Remove(row);
            UpdateNumbers();
        }

        /// <summary>
        /// Creates slides collection.
        /// </summary>
        /// <returns></returns>
        public static SlideCollection Create(PresentationPart presentationPart, PresentationSc presentation)
        {

            var numSlides = presentationPart.SlideParts.Count();
            var slideCollection = new List<SlideSc>(numSlides);
            var sldNumDic = new Dictionary<SlideSc, SlideNumber>(numSlides);
            for (var sldIndex = 0; sldIndex < numSlides; sldIndex++)
            {
                SlidePart slidePart = presentationPart.GetSlidePartByIndex(sldIndex);
                var slideNumber = new SlideNumber(sldIndex + 1);
                var newSlide = new SlideSc(slidePart, slideNumber, presentation);
                sldNumDic.Add(newSlide, slideNumber);
                slideCollection.Add(newSlide);
            }

            return new SlideCollection(slideCollection, presentationPart, sldNumDic);
        }

        #region Private Methods

        private void RemoveFromDom(in int number)
        {
            P.Presentation presentation = _presentationPart.Presentation;
            // gets the list of slide identifiers in the presentation
            SlideIdList slideIdList = presentation.SlideIdList;
            // gets the slide identifier of the specified slide
            SlideId slideId = (SlideId)slideIdList.ChildElements[number - 1];
            // gets the relationship identifier of the slide
            string slideRelId = slideId.RelationshipId;
            // removes the slide from the slide list
            slideIdList.RemoveChild(slideId);

            // remove references to the slide from all custom shows
            if (presentation.CustomShowList != null)
            {
                // iterates through the list of custom shows
                foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList == null)
                    {
                        continue;
                    }

                    // declares a link list of slide list entries
                    var slideListEntries = new LinkedList<SlideListEntry>();
                    foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                    {
                        // finds the slide reference to remove from the custom show
                        if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                        {
                            slideListEntries.AddLast(slideListEntry);
                        }
                    }
                    // removes all references to the slide from the custom show
                    foreach (SlideListEntry slideListEntry in slideListEntries)
                    {
                        customShow.SlideList.RemoveChild(slideListEntry);
                    }
                }
            }

            // gets the slide part for the specified slide
            SlidePart slidePart = _presentationPart.GetPartById(slideRelId) as SlidePart;
            // removes the slide part
            _presentationPart.DeletePart(slidePart);
        }

        private void UpdateNumbers()
        {
            var current = 0;
            foreach (SlideSc slide in CollectionItems)
            {
                current++;
                _slideToSlideNumber[slide].Number = current;
            }
        }

        #endregion Private Methods
    }
}