using System;
using System.Collections.Generic;
using SlideXML.Models.SlideComponents;
using SlideXML.Validation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
namespace SlideXML.Models.Settings
{
    /// <summary>
    /// Represents presentation settings.
    /// </summary>
    public class PreSettings : IPreSettings
    {
        private readonly Lazy<Dictionary<int, int>> _lvlFontHeights;

        #region Properties

        /// <summary>
        /// Gets default level sizes.
        /// </summary>
        /// <returns></returns>
        public Dictionary<int, int> LlvFontHeights => _lvlFontHeights.Value;

        #endregion Properties

        #region Constructors

        public PreSettings(P.Presentation xmlPresentation)
        {
            Check.NotNull(xmlPresentation, nameof(xmlPresentation));
            _lvlFontHeights = new Lazy<Dictionary<int, int>>(ParseLlvSizes(xmlPresentation));
        }

        #endregion Constructors

        private static Dictionary<int, int> ParseLlvSizes(P.Presentation xmlPresentation)
        {
            var levelSizes = new Dictionary<int, int>();
            foreach (var textPr in xmlPresentation.DefaultTextStyle.Elements<A.TextParagraphPropertiesType>())
            {
                var fs = textPr.GetFirstChild<A.DefaultRunProperties>()?.FontSize;
                if (fs == null)
                {
                    continue;
                }
                // fourth character of LocalName contains level number, example: "lvl1pPr, lvl2pPr, etc."
                var lvl = int.Parse(textPr.LocalName[3].ToString());
                levelSizes.Add(lvl, fs.Value);
            }

            return levelSizes;
        }
    }
}
