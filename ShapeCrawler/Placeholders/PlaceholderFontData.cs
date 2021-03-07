using System.Collections.Generic;

namespace ShapeCrawler.Placeholders
{
    /// <summary>
    ///     Represents placeholder font data.
    /// </summary>
    internal class PlaceholderFontData : PlaceholderData
    {
        #region Properties

        internal Dictionary<int, FontData> LvlToFontData;

        #endregion Properties
    }
}