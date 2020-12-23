using System.Collections.Generic;

namespace ShapeCrawler.Services.Placeholders
{
    /// <summary>
    /// Represents placeholder font data.
    /// </summary>
    /// TODO: consider to union PlaceholderData, PlaceholderLocationData, PlaceholderFontData
    public class PlaceholderFontData : PlaceholderData
    {
        #region Properties

        public Dictionary<int, int> LvlFontHeights;

        #endregion Properties
    }
}