using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Services;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Models.Settings
{
    /// <summary>
    /// <inheritdoc cref="IPreSettings"/>
    /// </summary>
    public class PreSettings : IPreSettings
    {
        private readonly Lazy<Dictionary<int, int>> _lvlFontHeights;

        #region Properties

        /// <summary>
        /// <inheritdoc cref="IPreSettings.LlvFontHeights"/>
        /// </summary>
        public Dictionary<int, int> LlvFontHeights => _lvlFontHeights.Value;

        /// <summary>
        /// <inheritdoc cref="IPreSettings.XlsxDocuments"/>
        /// </summary>
        public Dictionary<OpenXmlPart, SpreadsheetDocument> XlsxDocuments { get; }

        public Lazy<SlideSize> SlideSize { get; }

        #endregion Properties

        #region Constructors

        public PreSettings(P.Presentation sdkPresentation, Lazy<SlideSize> slideSize)
        {
            Check.NotNull(sdkPresentation, nameof(sdkPresentation));

            SlideSize = slideSize ?? throw new ArgumentNullException(nameof(slideSize));
            _lvlFontHeights = new Lazy<Dictionary<int, int>>(()=>ParseFontHeights(sdkPresentation));
            XlsxDocuments = new Dictionary<OpenXmlPart, SpreadsheetDocument>(); //TODO: make lazy initialization
        }

        #endregion Constructors

        #region Private Methods

        private static Dictionary<int, int> ParseFontHeights(P.Presentation xmlPresentation)
        {
            var lvlFontHeights = new Dictionary<int, int>();

            // from presentation default text settings
            if (xmlPresentation.DefaultTextStyle != null)
            {
                lvlFontHeights = FontHeightParser.FromCompositeElement(xmlPresentation.DefaultTextStyle);
            }

            // from theme default text settings
            if (!lvlFontHeights.Any())
            {
                var themeTextDefault = xmlPresentation.PresentationPart.ThemePart.Theme.ObjectDefaults.TextDefault;
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
