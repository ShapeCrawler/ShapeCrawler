﻿using System.Collections.Generic;
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
    /// Represents a collection of the slides.
    /// </summary>
    public class SlideCollection : EditableCollection<SlideSc>
    {
        #region Fields

        private readonly PresentationPart _sdkPrePart;
        private readonly Dictionary<SlideSc, SlideNumber> _sldNumDic;

        #endregion Fields

        #region Constructors

        private SlideCollection(List<SlideSc> slides, PresentationPart sdkPrePart, Dictionary<SlideSc, SlideNumber> sldNumDic)
        {
            CollectionItems = slides;
            _sdkPrePart = sdkPrePart;
            _sldNumDic = sldNumDic;
        }

        #endregion Constructors

        /// <summary>
        /// Removes the specified slide.
        /// </summary>
        /// <param name="item"></param>
        public override void Remove(SlideSc item)
        {
            Check.NotNull(item, nameof(item));

            RemoveFromDom(item.Number);
            _sdkPrePart.Presentation.Save();
            CollectionItems.Remove(item);
            UpdateNumbers();
        }

        /// <summary>
        /// Creates slides collection.
        /// </summary>
        /// <returns></returns>
        public static SlideCollection Create(PresentationPart sdkPrePart, IPresentationData preSettings, PresentationSc presentationEx)
        {
            Check.NotNull(sdkPrePart, nameof(sdkPrePart));
            Check.NotNull(preSettings, nameof(preSettings));

            var numSlides = sdkPrePart.SlideParts.Count();
            var slideCollection = new List<SlideSc>(numSlides);
            var sldNumDic = new Dictionary<SlideSc, SlideNumber>(numSlides);
            for (var sldIndex = 0; sldIndex < numSlides; sldIndex++)
            {
                var sdkSldPart = sdkPrePart.GetSlidePartByIndex(sldIndex);
                var sldNumEntity = new SlideNumber(sldIndex + 1);
                var newSlide = new SlideSc(sdkSldPart, sldNumEntity, preSettings, presentationEx);
                sldNumDic.Add(newSlide, sldNumEntity);
                slideCollection.Add(newSlide);
            }

            return new SlideCollection(slideCollection, sdkPrePart, sldNumDic);
        }

        #region Private Methods

        private void RemoveFromDom(int number)
        {
            P.Presentation presentation = _sdkPrePart.Presentation;
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
            SlidePart slidePart = _sdkPrePart.GetPartById(slideRelId) as SlidePart;
            // removes the slide part
            _sdkPrePart.DeletePart(slidePart);
        }

        private void UpdateNumbers()
        {
            var current = 0;
            foreach (var slide in CollectionItems)
            {
                current++;
                _sldNumDic[slide].Number = current;
            }
        }

        #endregion Private Methods
    }
}