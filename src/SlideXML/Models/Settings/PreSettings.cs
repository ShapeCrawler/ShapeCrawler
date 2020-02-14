using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
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

        #region Private Methods

        private static Dictionary<int, int> ParseLlvSizes(P.Presentation xmlPresentation)
        {
            var levelSizes = new Dictionary<int, int>();
            
            FromCompositeElement(xmlPresentation.DefaultTextStyle); // from presentation default text settings
            if (!levelSizes.Any())
            {
                // parses from theme default text settings
                FromCompositeElement(xmlPresentation.PresentationPart.ThemePart.Theme.ObjectDefaults.TextDefault.ListStyle);
            }

            // local function
            void FromCompositeElement(OpenXmlCompositeElement ce)
            {
                foreach (var textPr in ce.Elements().Where(e => e.LocalName.StartsWith("lvl"))) // <a:lvl1pPr>, <a:lvl2pPr>, etc.
                {
                    var fs = textPr.GetFirstChild<A.DefaultRunProperties>()?.FontSize;
                    if (fs == null)
                    {
                        continue;
                    }
                    // fourth character of LocalName contains level number, example: "lvl1pPr -> 1, lvl2pPr -> 2, etc."
                    var lvl = int.Parse(textPr.LocalName[3].ToString());
                    levelSizes.Add(lvl, fs.Value);
                }
            }

            return levelSizes;
        }

        #endregion Private Methods
    }
}
