using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace SlideDotNet.Models.Settings
{
    /// <summary>
    /// Represents a global presentation settings.
    /// </summary>
    public interface IPreSettings
    {
        /// <summary>
        /// Returns font heights from global presentation or theme settings.
        /// </summary>
        Dictionary<int, int> LlvFontHeights { get; }

        /// <summary>
        /// Returns cache Excel documents instantiated by chart shapes.
        /// </summary>
        public Dictionary<OpenXmlPart, SpreadsheetDocument> XlsxDocuments { get; }

        public Lazy<SlideSize> SlideSize { get; }
    }
}