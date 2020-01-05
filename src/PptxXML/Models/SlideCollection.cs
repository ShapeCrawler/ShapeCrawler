using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ObjectEx.Utilities;
using System.Collections;
using System.Collections.Generic;
using P = DocumentFormat.OpenXml.Presentation;

namespace PptxXML.Models
{
    /// <summary>
    /// Represents a collection of a slides.
    /// </summary>
    public class SlideCollection: ISlideCollection
    {
        #region Fields

        private readonly List<SlideEx> _items;
        private readonly PresentationDocument _xmlPreDoc;

        #endregion Fields

        #region Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="SlideCollection"></see> class.
        /// </summary>
        public SlideCollection(PresentationDocument xmlPreDoc)
        {
            Check.NotNull(xmlPreDoc, nameof(xmlPreDoc));
            _items = new List<SlideEx>();
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="SlideCollection"></see> class.
        /// </summary>
        public SlideCollection(PresentationDocument xmlPreDoc, int sldNumbers)
        {
            Check.NotNull(xmlPreDoc, nameof(xmlPreDoc));
            Check.IsPositive(sldNumbers, nameof(sldNumbers));
            _xmlPreDoc = xmlPreDoc;
            _items = new List<SlideEx>(sldNumbers);
        }

        #endregion Constructors

        #region Public Methods

        /// <summary>
        /// Adds slide item.
        /// </summary>
        /// <param name="item"></param>
        public void Add(SlideEx item)
        {
            Check.NotNull(item, nameof(item));
            _items.Add(item);
        }

        /// <summary>
        /// Removes specified slide and saves DOM.
        /// </summary>
        public void Remove(SlideEx item)
        {
            Check.NotNull(item, nameof(item));

            RemoveFromDom(item.Number);
            _xmlPreDoc.PresentationPart.Presentation.Save(); // save the modified presentation

            _items.Remove(item);
            UpdateNumbers();
        }

        /// <summary>
        /// Returns an enumerator for slide list.
        /// </summary>
        /// <returns></returns>
        public IEnumerator<SlideEx> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        /// <summary>
        /// Gets item by index.
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public SlideEx this[int index] => _items[index];

        #endregion Public Methods

        #region Private Methods

        private void RemoveFromDom(int number)
        {
            PresentationPart presentationPart = _xmlPreDoc.PresentationPart;

            // Get the presentation from the presentation part.
            P.Presentation presentation = presentationPart.Presentation;

            // Get the list of slide IDs in the presentation.
            SlideIdList slideIdList = presentation.SlideIdList;

            // Get the slide ID of the specified slide
            SlideId slideId = slideIdList.ChildElements[number - 1] as SlideId;

            // Get the relationship ID of the slide.
            string slideRelId = slideId.RelationshipId;

            // Remove the slide from the slide list.
            slideIdList.RemoveChild(slideId);

            // Remove references to the slide from all custom shows.
            if (presentation.CustomShowList != null)
            {
                // Iterate through the list of custom shows.
                foreach (var customShow in presentation.CustomShowList.Elements<CustomShow>())
                {
                    if (customShow.SlideList != null)
                    {
                        // Declare a link list of slide list entries.
                        LinkedList<SlideListEntry> slideListEntries = new LinkedList<SlideListEntry>();
                        foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                        {
                            // Find the slide reference to remove from the custom show.
                            if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                            {
                                slideListEntries.AddLast(slideListEntry);
                            }
                        }

                        // Remove all references to the slide from the custom show.
                        foreach (SlideListEntry slideListEntry in slideListEntries)
                        {
                            customShow.SlideList.RemoveChild(slideListEntry);
                        }
                    }
                }
            }

            // Get the slide part for the specified slide.
            SlidePart slidePart = presentationPart.GetPartById(slideRelId) as SlidePart;

            // Remove the slide part.
            presentationPart.DeletePart(slidePart);
        }

        private void UpdateNumbers()
        {
            var current = 0;
            foreach (var i in _items)
            {
                i.Number = ++current;
            }
        }

        #endregion Private Methods
    }
}
