using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Office2013.Word;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "Exception")]
    internal class SCSection : ISection
    {
        internal readonly SCPresentation presentation;

        internal SCSection(SCPresentation presentation, P14.Section p14Section)
        {
            this.presentation = presentation;
            this.SDKSection = p14Section;
        }

        public ISectionSlideCollection Slides => SCSectionSlideCollection.Create(this);

        public Section SDKSection { get; }

        public string Name => GetName();

        private string GetName()
        {
            return this.SDKSection.Name;
        }
    }

    internal class SCSectionSlideCollection : ISectionSlideCollection
    {
        private readonly SCSection section;
        private readonly List<SCSlide> slides;
        private readonly Dictionary<SCSlide, SectionSlideIdListEntry> slideToSectionSlideIdListEntryDic;

        private SCSectionSlideCollection(
            SCSection section,
            List<SCSlide> sectionSlides,
            Dictionary<SCSlide, SectionSlideIdListEntry> slideToSectionSlideIdListEntryDic)
        {
            this.section = section;
            this.slides = sectionSlides;
            this.slideToSectionSlideIdListEntryDic = slideToSectionSlideIdListEntryDic;
        }

        internal static ISectionSlideCollection Create(SCSection section)
        {
            var sectionSlides = new List<SCSlide>();
            var slideToSectionSlideIdListEntry = new Dictionary<SCSlide, SectionSlideIdListEntry>();
            foreach (var sectionSlideIdListEntry in section.SDKSection.Descendants<P14.SectionSlideIdListEntry>())
            {
                var slide = section.presentation.SlidesInternal.GetBySlideId(sectionSlideIdListEntry.Id);
                sectionSlides.Add(slide);
                slideToSectionSlideIdListEntry.Add(slide, sectionSlideIdListEntry);
            }

            return new SCSectionSlideCollection(section, sectionSlides, slideToSectionSlideIdListEntry);
        }

        public IEnumerator<ISlide> GetEnumerator()
        {
            return this.slides.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return this.GetEnumerator();
        }

        public int Count => this.slides.Count;

        public ISlide this[int index] => this.slides[index];

        public void Remove(ISlide slide)
        {
            var slideInternal = (SCSlide)slide;

            this.slides.Remove(slideInternal);
            var exist = this.slideToSectionSlideIdListEntryDic.TryGetValue(slideInternal, out var sectionSlideIdListEntry);
            if (exist)
            {
                sectionSlideIdListEntry.Remove();
            }
        }
    }

    public interface ISectionSlideCollection : IEnumerable<ISlide>
    {
        int Count { get; }

        ISlide this[int index] { get; }
        void Remove(ISlide slide);
    }
}