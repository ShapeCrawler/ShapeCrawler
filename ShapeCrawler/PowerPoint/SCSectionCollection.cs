using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler
{
    /// <summary>
    /// <inheritdoc cref="ISectionCollection"/>
    /// </summary>
    internal class SCSectionCollection : ISectionCollection
    {
        private readonly List<SCSection> sections;
        private readonly SectionList? sdkSectionList;

        private SCSectionCollection(List<SCSection> sections)
        {
            this.sections = sections;
        }

        private SCSectionCollection(List<SCSection> sections, SectionList sdkSectionList)
        {
            this.sections = sections;
            this.sdkSectionList = sdkSectionList;
        }

        public int Count => this.sections.Count;

        public ISection this[int i] => this.sections[i];

        public static ISectionCollection Create(SCPresentation presentation)
        {
            var sections = new List<SCSection>();
            var sectionList = presentation.PresentationDocument.PresentationPart!.Presentation.PresentationExtensionList?.Descendants<SectionList>().FirstOrDefault();

            if (sectionList == null)
            {
                return new SCSectionCollection(sections);
            }

            foreach (var sectionXml in sectionList)
            {
                var sdkSection = (Section)sectionXml;
                var sectionSlides = new List<SCSlide>();
                foreach (var slideId in sdkSection.Descendants<SlideId>())
                {
                    var slide = presentation.SlidesInternal.GetBySlideId(slideId);
                    sectionSlides.Add(slide);
                }

                sections.Add(new SCSection(sectionSlides, sdkSection));
            }

            return new SCSectionCollection(sections, sectionList);
        }


        public IEnumerator<ISection> GetEnumerator()
        {
            return this.sections.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.sections.GetEnumerator();
        }


        public void Remove(ISection removingSection)
        {
            if (this.Count == 0)
            {
                return;
            }

            ((SCSection)removingSection).SDKSection.Remove();

            if (this.sections.Count == 1)
            {
                this.sdkSectionList.Remove();
            }
        }

        public ISection GetByName(string sectionName)
        {
            return this.sections.First(section => section.Name == sectionName);
        }
    }
}