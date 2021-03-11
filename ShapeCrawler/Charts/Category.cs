using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler.Shared;
using System;
using System.Collections.Generic;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart category.
    /// </summary>
    public class Category
    {
        private int _xCatIdx;
        private readonly NumericValue _cachedCatName;
        private ResettableLazy<Dictionary<int, X.Cell>> _catIdxToXCell;

        #region Constructors

        internal Category(ResettableLazy<Dictionary<int, X.Cell>> catIdxToXCell, int xCatIdx, NumericValue cachedCatName, Category mainCategory)
            :this(catIdxToXCell, xCatIdx, cachedCatName)
        {
            // TODO: what about creating a new separate class like MultiCategory:Category
            MainCategory = mainCategory;
        }

        internal Category(ResettableLazy<Dictionary<int, X.Cell>> catIdxToXCell, int xCatIdx, NumericValue cachedCatName)
        {
            _catIdxToXCell = catIdxToXCell;
            _xCatIdx = xCatIdx;
            Name = cachedCatName.InnerText;
        }

        #endregion Constructors

        #region Properties

        /// <summary>
        ///     Gets main category. Returns <c>NULL</c> if the chart is not a multi-category chart type.
        /// </summary>
        public Category MainCategory { get; }

#if DEBUG
        /// <summary>
        ///     Gets or sets category name.
        /// </summary>
        public string Name { get; set; }
#else
        /// <summary>
        ///     Gets category name.
        /// </summary>
        public string Name { get; }
#endif

        #endregion Properties
    }
}