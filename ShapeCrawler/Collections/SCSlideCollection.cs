using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Collections
{
    internal class SCSlideCollection : ISlideCollection
    {
        private readonly SCPresentation parentPresentation;
        private readonly ResettableLazy<List<SCSlide>> slides;
        private PresentationPart presentationPart;

        internal SCSlideCollection(SCPresentation presentation)
        {
            this.presentationPart = presentation.PresentationDocument.PresentationPart ?? throw new ArgumentNullException("PresentationPart");
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
            // TODO: slide layout and master of removed slide also should be deleted if they are unused
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

        public void Add(ISlide outerSlide)
        {
            SCSlide outerInnerSlide = (SCSlide)outerSlide;
            if (outerInnerSlide.ParentPresentation == this.parentPresentation)
            {
                throw new ShapeCrawlerException("Adding slide cannot be belong to the same presentation.");
            }

            this.parentPresentation.ThrowIfClosed();

            var presentation = (SCPresentation)outerInnerSlide.ParentPresentation;
            PresentationDocument addingSlideDoc = presentation.PresentationDocument;
            PresentationDocument destDoc = this.parentPresentation.PresentationDocument;
            PresentationPart addingPresentationPart = addingSlideDoc.PresentationPart;
            PresentationPart destPresentationPart = destDoc.PresentationPart;
            Presentation destPresentation = destPresentationPart.Presentation;
            int addingSlideIndex = outerSlide.Number - 1;
            SlideId addingSlideId = (SlideId)addingPresentationPart.Presentation.SlideIdList.ChildElements[addingSlideIndex];
            SlidePart addingSlidePart = (SlidePart)addingPresentationPart.GetPartById(addingSlideId.RelationshipId);

            SlidePart addedSlidePart = destPresentationPart.AddPart(addingSlidePart);
            NotesSlidePart noticePart = addedSlidePart.GetPartsOfType<NotesSlidePart>().FirstOrDefault();
            if (noticePart != null)
            {
                addedSlidePart.DeletePart(noticePart);
            }
            SlideMasterPart addedSlideMasterPart = destPresentationPart.AddPart(addedSlidePart.SlideLayoutPart.SlideMasterPart);

            // Create new slide ID
            SlideId slideId = new ()
            {
                Id = CreateId(destPresentation.SlideIdList),
                RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlidePart)
            };
            destPresentation.SlideIdList.Append(slideId);

            // Create new master slide ID
            uint masterId = CreateId(destPresentation.SlideMasterIdList);
            SlideMasterId slideMaterId = new ()
            {
                Id = masterId,
                RelationshipId = destDoc.PresentationPart.GetIdOfPart(addedSlideMasterPart)
            };
            destDoc.PresentationPart.Presentation.SlideMasterIdList.Append(slideMaterId);

            destDoc.PresentationPart.Presentation.Save();

            // Make sure that all slide layouts have unique ids.
            foreach (SlideMasterPart slideMasterPart in destDoc.PresentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                {
                    masterId++;
                    slideLayoutId.Id = masterId;
                }

                slideMasterPart.SlideMaster.Save();
            }

            this.slides.Reset();
            this.parentPresentation.SlideMastersValue.Reset();
        }

        public void Insert(int position, ISlide outerSlide)
        {
            if (position < 1 || position > this.slides.Value.Count + 1)
            {
                throw new ArgumentOutOfRangeException(nameof(position));
            }

            this.Add(outerSlide);
            int addedSlideIndex = this.slides.Value.Count - 1;
            this.slides.Value[addedSlideIndex].Number = position;

            this.slides.Reset();
            this.parentPresentation.SlideMastersValue.Reset();
        }

        private static uint CreateId(SlideIdList slideIdList)
        {
            uint currentId = 0;
            foreach (SlideId slideId in slideIdList)
            {
                if (slideId.Id > currentId)
                {
                    currentId = slideId.Id;
                }
            }

            return ++currentId;
        }

        private static uint CreateId(SlideMasterIdList slideMasterIdList)
        {
            uint currentId = 0;
            foreach (SlideMasterId masterId in slideMasterIdList)
            {
                if (masterId.Id > currentId)
                {
                    currentId = masterId.Id;
                }
            }

            return ++currentId;
        }

        private List<SCSlide> GetSlides()
        {
            this.presentationPart = this.parentPresentation.PresentationDocument.PresentationPart!;
            int slidesCount = this.presentationPart.SlideParts.Count();
            var slides = new List<SCSlide>(slidesCount);
            var slideIds = this.presentationPart.Presentation.SlideIdList.ChildElements.OfType<SlideId>().ToList();
            for (var slideIndex = 0; slideIndex < slidesCount; slideIndex++)
            {
                var slideId = slideIds[slideIndex];
                var slidePart = (SlidePart)this.presentationPart.GetPartById(slideId.RelationshipId);
                var newSlide = new SCSlide(this.parentPresentation, slidePart, slideId);
                slides.Add(newSlide);
            }

            return slides;
        }

        internal SCSlide GetBySlideId(string slideId)
        {
            return this.slides.Value.First(scSlide => scSlide.SlideId.Id == slideId);
        }
    }
}