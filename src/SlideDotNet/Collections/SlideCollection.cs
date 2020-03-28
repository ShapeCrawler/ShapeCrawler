using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideDotNet.Extensions;
using SlideDotNet.Models;
using SlideDotNet.Models.Settings;
using SlideDotNet.Validation;
using Slide = SlideDotNet.Models.Slide;

namespace SlideDotNet.Collections
{
    /// <summary>
    /// Represents a collection of the slides.
    /// </summary>
    public class SlideCollection : EditAbleCollection<Slide>
    {
        #region Fields

        private readonly PresentationDocument _sdkPre;
        private readonly Dictionary<Slide, SlideNumber> _sldNumEntities;

        #endregion Fields

        #region Constructors

        private SlideCollection(List<Slide> slides, PresentationDocument sdkPre, Dictionary<Slide, SlideNumber> sldNumEntities)
        {
            CollectionItems = slides;
            _sdkPre = sdkPre;
            _sldNumEntities = sldNumEntities;
        }

        #endregion Constructors

        /// <summary>
        /// Removes the specified slide.
        /// </summary>
        /// <param name="item"></param>
        public override void Remove(Slide item)
        {
            Check.NotNull(item, nameof(item));

            RemoveFromDom(item.Number);
            _sdkPre.PresentationPart.Presentation.Save();
            CollectionItems.Remove(item);
            UpdateNumbers();
        }

        /// <summary>
        /// Creates slides collection.
        /// </summary>
        /// <param name="xmlDoc"></param>
        /// <param name="preSettings"></param>
        /// <returns></returns>
        public static SlideCollection Create(PresentationDocument xmlDoc, IPreSettings preSettings)
        {
            var xmlPrePart = xmlDoc.PresentationPart;
            var slideCollection = new List<Slide>();
            var sldNumDic = new Dictionary<Slide, SlideNumber>();
            for (var sldIndex = 0; sldIndex < xmlPrePart.SlideParts.Count(); sldIndex++)
            {
                var xmlSldPart = xmlPrePart.GetSlidePartByIndex(sldIndex);
                var sldNumEntity = new SlideNumber(sldIndex + 1);
                var newSlide = new Slide(xmlSldPart, sldNumEntity, preSettings);
                sldNumDic.Add(newSlide, sldNumEntity);
                slideCollection.Add(newSlide);
            }

            return new SlideCollection(slideCollection, xmlDoc, sldNumDic);
        }

        #region Private Methods

        private void RemoveFromDom(int number)
        {
            PresentationPart presentationPart = _sdkPre.PresentationPart;
            DocumentFormat.OpenXml.Presentation.Presentation presentation = presentationPart.Presentation;
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
            SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;
            // removes the slide part
            presentationPart.DeletePart(slidePart);
        }

        private void UpdateNumbers()
        {
            var current = 0;
            foreach (var slide in CollectionItems)
            {
                current++;
                _sldNumEntities[slide].Number = current;
            }
        }

        #endregion Private Methods
    }
}