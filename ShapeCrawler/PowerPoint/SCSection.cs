﻿using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace ShapeCrawler
{
    [SuppressMessage("ReSharper", "InconsistentNaming", Justification = "Exception")]
    internal class SCSection : ISection
    {
        internal readonly SCSectionCollection Sections;

        internal SCSection(SCSectionCollection sections, P14.Section p14Section)
        {
            this.Sections = sections;
            this.SDKSection = p14Section;
        }

        public ISectionSlideCollection Slides => new SCSectionSlideCollection(this);

        public Section SDKSection { get; }

        public string Name => GetName();

        private string GetName()
        {
            return this.SDKSection.Name;
        }
    }
}