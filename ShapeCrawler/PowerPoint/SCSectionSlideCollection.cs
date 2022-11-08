using System;
using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler;

internal class SCSectionSlideCollection : ISectionSlideCollection
{
    private readonly SCSection parentSection;
    private List<SCSlide> sectionSlides;

    public SCSectionSlideCollection(SCSection parentSection)
    {
        this.parentSection = parentSection;
        var slides = parentSection.Sections.Presentation.SlidesInternal;
        slides.CollectionChanged += this.OnPresSlideCollectionChanged;

        this.sectionSlides = new List<SCSlide>();
        this.Initialize();
    }

    public int Count => this.sectionSlides.Count;

    public ISlide this[int index] => this.sectionSlides[index];

    public IEnumerator<ISlide> GetEnumerator()
    {
        return this.sectionSlides.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return this.GetEnumerator();
    }

    private void OnPresSlideCollectionChanged(object sender, EventArgs e)
    {
        this.Initialize();
    }

    private void Initialize()
    {
        this.sectionSlides = new List<SCSlide>();
        foreach (var sectionSlideIdListEntry in this.parentSection.SDKSection.Descendants<SectionSlideIdListEntry>())
        {
            var slide = this.parentSection.Sections.Presentation.SlidesInternal.GetBySlideId(sectionSlideIdListEntry.Id!);
            this.sectionSlides.Add(slide);
        }
    }
}