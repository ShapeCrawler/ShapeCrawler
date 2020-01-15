using System.Collections.Generic;
using ObjectEx.Utilities;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
namespace PptxXML.Models.Settings
{
    /// <summary>
    /// Represents presentation settings.
    /// </summary>
    public class PreSettings : IPreSettings
    {
        private readonly P.Presentation _xmlPresentation;
        private Dictionary<int, int> _levelSizes;

        #region Constructors

        public PreSettings(P.Presentation xmlPresentation)
        {
            Check.NotNull(xmlPresentation, nameof(xmlPresentation));
            _xmlPresentation = xmlPresentation;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets default level sizes.
        /// </summary>
        /// <returns></returns>
        public Dictionary<int, int> LlvFontHeights {
            get
            {
                if (_levelSizes == null)
                {
                    _levelSizes = new Dictionary<int, int>();
                    var defTxtStyle = _xmlPresentation.DefaultTextStyle;
                    _levelSizes.Add(1, defTxtStyle.Level1ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(2, defTxtStyle.Level2ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(3, defTxtStyle.Level3ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(4, defTxtStyle.Level4ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(5, defTxtStyle.Level5ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(6, defTxtStyle.Level6ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(7, defTxtStyle.Level7ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(8, defTxtStyle.Level8ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                    _levelSizes.Add(9, defTxtStyle.Level9ParagraphProperties.GetFirstChild<A.DefaultRunProperties>().FontSize.Value);
                }

                return _levelSizes;
            } 
        }

        #endregion
    }
}
