using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;
using SlideSize = ShapeCrawler.Models.SlideSize;

namespace ShapeCrawler.Settings
{
    /// <summary>
    /// <inheritdoc cref="IPresentationData"/>
    /// </summary>
    public class PresentationData : IPresentationData
    {
        private readonly Lazy<Dictionary<int, int>> _lvlFontHeights;

        #region Properties

        /// <summary>
        /// <inheritdoc cref="IPresentationData.LlvFontHeights"/>
        /// </summary>
        public Dictionary<int, int> LlvFontHeights => _lvlFontHeights.Value;

        /// <summary>
        /// <inheritdoc cref="IPresentationData.XlsxDocuments"/>
        /// </summary>
        public Dictionary<OpenXmlPart, SpreadsheetDocument> XlsxDocuments { get; }

        public Lazy<SlideSize> SlideSize { get; }

        #endregion Properties

        #region Constructors

        public PresentationData(P.Presentation sdkPresentation, Lazy<SlideSize> slideSize)
        {
            Check.NotNull(sdkPresentation, nameof(sdkPresentation));

            SlideSize = slideSize ?? throw new ArgumentNullException(nameof(slideSize));
            _lvlFontHeights = new Lazy<Dictionary<int, int>>(()=>ParseFontHeights(sdkPresentation));
            XlsxDocuments = new Dictionary<OpenXmlPart, SpreadsheetDocument>(); //TODO: make lazy initialization
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
