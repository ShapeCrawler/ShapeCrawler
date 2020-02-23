using System;
using System.Collections.Generic;
using System.Linq;
using SlideDotNet.Services;
using SlideDotNet.Validation;
using P = DocumentFormat.OpenXml.Presentation;

namespace SlideDotNet.Models.Settings
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

        #endregion Properties

        #region Constructors

        public PreSettings(P.Presentation xmlPresentation)
        {
            Check.NotNull(xmlPresentation, nameof(xmlPresentation));
            _lvlFontHeights = new Lazy<Dictionary<int, int>>(ParseFontHeights(xmlPresentation));
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
