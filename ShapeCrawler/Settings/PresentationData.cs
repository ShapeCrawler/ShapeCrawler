using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories;
using ShapeCrawler.Models;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Settings
{
    internal class PresentationData
    {
        private readonly Lazy<Dictionary<int, int>> _lvlFontHeights;

        #region Constructors

        public PresentationData(P.Presentation pPresentation, Lazy<SlideSizeSc> slideSize)
        {
            SlideSize = slideSize ?? throw new ArgumentNullException(nameof(slideSize));
            _lvlFontHeights = new Lazy<Dictionary<int, int>>(() => ParseFontHeights(pPresentation));
            SpreadsheetCache = new Dictionary<EmbeddedPackagePart, SpreadsheetDocument>();
        }

        #endregion Constructors

        #region Private Methods

        private static Dictionary<int, int> ParseFontHeights(P.Presentation pPresentation)
        {
            var lvlFontHeights = new Dictionary<int, int>();

            // from presentation default text settings
            if (pPresentation.DefaultTextStyle != null)
            {
                lvlFontHeights = FontHeightParser.FromCompositeElement(pPresentation.DefaultTextStyle);
            }

            // from theme default text settings
            if (!lvlFontHeights.Any())
            {
                var themeTextDefault = pPresentation.PresentationPart.ThemePart.Theme.ObjectDefaults.TextDefault;
                if (themeTextDefault != null)
                {
                    lvlFontHeights = FontHeightParser.FromCompositeElement(themeTextDefault.ListStyle);
                }
            }

            return lvlFontHeights;
        }

        #endregion Private Methods

        #region Properties

        public Dictionary<int, int> LlvFontHeights => _lvlFontHeights.Value;

        /// <summary>
        ///     Returns cache Excel documents instantiated by chart shapes.
        /// </summary>
        public Dictionary<EmbeddedPackagePart, SpreadsheetDocument>
            SpreadsheetCache { get; } //TODO: move it up to Presentation level

        public Lazy<SlideSizeSc> SlideSize { get; }

        #endregion Properties
    }
}