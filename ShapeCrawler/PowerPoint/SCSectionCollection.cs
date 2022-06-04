using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler
{
    /// <summary>
    /// <inheritdoc cref="ISectionCollection"/>
    /// </summary>
    internal class SCSectionCollection : ISectionCollection
    {
        private readonly List<SCSection> sections;
        private readonly SectionList? sdkSectionList;
        private readonly SCPresentation presentation;

        private SCSectionCollection(SCPresentation presentation, List<SCSection> sections)
        {
            this.presentation = presentation;
            this.sections = sections;
        }

        private SCSectionCollection(SCPresentation presentation, List<SCSection> sections, SectionList sdkSectionList)
        {
            this.presentation = presentation;
            this.sections = sections;
            this.sdkSectionList = sdkSectionList;
        }

        public int Count => this.sections.Count;

        public ISection this[int i] => this.sections[i];

        internal static SCSectionCollection Create(SCPresentation presentation)
        {
            var sections = new List<SCSection>();
            var sdkSectionList = presentation.PresentationDocument.PresentationPart!.Presentation.PresentationExtensionList?.Descendants<SectionList>().FirstOrDefault();

            if (sdkSectionList == null)
            {
                return new SCSectionCollection(presentation, sections);
            }

            foreach (P14.Section p14Section in sdkSectionList)
            {
                sections.Add(new SCSection(presentation, p14Section));
            }

            return new SCSectionCollection(presentation, sections, sdkSectionList);
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
            if (this.sdkSectionList == null || this.Count == 0)
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

        internal void RemoveSldId(string id)
        {
            var removing = this.sdkSectionList?.Descendants<P14.SectionSlideIdListEntry>().FirstOrDefault(s => s.Id == id);
            if (removing == null)
            {
                return;
            }

            removing.Remove();
            this.presentation.PresentationDocument.PresentationPart.Presentation.Save();
        }
    }
}