using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Factories;
using ShapeCrawler.Models;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Settings
{
    public class PresentationData
    {
        private readonly Lazy<Dictionary<int, int>> _lvlFontHeights;

        #region Properties

        public Dictionary<int, int> LlvFontHeights => _lvlFontHeights.Value;

        /// <summary>
        /// Returns cache Excel documents instantiated by chart shapes.
        /// </summary>
        public Dictionary<OpenXmlPart, SpreadsheetDocument> XlsxDocuments { get; }

        public Lazy<SlideSizeSc> SlideSize { get; }

        #endregion Properties

        #region Constructors

        public PresentationData(P.Presentation pPresentation, Lazy<SlideSizeSc> slideSize)
        {
            Check.NotNull(pPresentation, nameof(pPresentation));

            SlideSize = slideSize ?? throw new ArgumentNullException(nameof(slideSize));
            _lvlFontHeights = new Lazy<Dictionary<int, int>>(() => ParseFontHeights(pPresentation));
            XlsxDocuments = new Dictionary<OpenXmlPart, SpreadsheetDocument>();
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
    }
}
