using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using SlideDotNet.Extensions;
using SlideDotNet.Models.Settings;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Models
{
    /// <summary>
    /// <inheritdoc cref="ISlideCollection"/>
    /// </summary>
    public class SlideCollection: ISlideCollection
    {
        #region Fields

        private readonly List<Slide> _slides;
        private readonly PresentationDocument _xmlDoc;
        private readonly Dictionary<Slide, SlideNumber> _sldNumEntities;

        #endregion Fields

        #region Constructors

        private SlideCollection(List<Slide> slides, PresentationDocument xmlDoc, Dictionary<Slide, SlideNumber> sldNumEntities)
        {
            _slides = slides;
            _xmlDoc = xmlDoc;
            _sldNumEntities = sldNumEntities;
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// <inheritdoc cref="ISlideCollection.Remove"/>
        /// </summary>
        public void Remove(Slide item)
        {
            //TODO: validate case when last slide is deleted
            Check.NotNull(item, nameof(item));

            RemoveFromDom(item.Number);
            _xmlDoc.PresentationPart.Presentation.Save(); // save the modified presentation
            _slides.Remove(item);
            UpdateNumbers();
        }

        /// <summary>
        /// Returns the element at the specified index.
        /// </summary>
        public Slide this[int index] => _slides[index];

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

        /// <summary>
        /// Returns an enumerator for slide list.
        /// </summary>
        public IEnumerator<Slide> GetEnumerator()
        {
            return _slides.GetEnumerator();
        }

        
        /// <summary>
        /// Returns an enumerator for slide list. 
        /// </summary>
        /// <returns></returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            //TODO: why two GetEnumerator() methods?
            return _slides.GetEnumerator();
        }

        #endregion Public Methods

        #region Private Methods

        private void RemoveFromDom(int number)
        {
            PresentationPart presentationPart = _xmlDoc.PresentationPart;
            P.Presentation presentation = presentationPart.Presentation;
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
            foreach (var slide in _slides)
            {
                current++;
                _sldNumEntities[slide].Number = current;
            }
        }

        #endregion Private Methods
    }

    public class SlideNumber
    {
        /// <summary>
        /// Gets or sets slide number.
        /// </summary>
        public int Number { get; set; }

        public SlideNumber(int sldNum)
        {
            Check.IsPositive(sldNum, nameof(sldNum));
            Number = sldNum;
        }
    }
}
