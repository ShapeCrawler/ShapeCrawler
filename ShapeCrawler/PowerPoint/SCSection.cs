using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "Exception")]
    internal class SCSection : ISection
    {
        internal readonly Section SDKSection;

        private readonly List<SCSlide> _sectionSlides;

        public SCSection(List<SCSlide> sectionSlides, Section sdkSection)
        {
            this._sectionSlides = sectionSlides;
            this.SDKSection = sdkSection;
        }

        public List<ISlide> Slides => this._sectionSlides.OfType<ISlide>().ToList();

        public string Name => GetName();

        private string GetName()
        {
            return this.SDKSection.Name;
        }
    }
}