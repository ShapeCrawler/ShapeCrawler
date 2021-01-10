using System.Collections.Generic;

namespace ShapeCrawler.Factories.Placeholders
{
    /// <summary>
    /// Represents placeholder font data.
    /// </summary>
    /// TODO: consider to union PlaceholderData, PlaceholderLocationData, PlaceholderFontData
    public class PlaceholderFontData : PlaceholderData
    {
        #region Properties

        internal Dictionary<int, int> LvlFontHeights;

        #endregion Properties
    }
}