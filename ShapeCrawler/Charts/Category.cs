using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing.Charts;
using ShapeCrawler.Shared;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts
{
    /// <summary>
    ///     Represents a chart category.
    /// </summary>
    public class Category
    {
        private int _catIdx;
        private NumericValue _cachedCatName;
        private ResettableLazy<List<X.Cell>> _catIdxToXCell;

        #region Constructors

        internal Category(
            ResettableLazy<List<X.Cell>> catIdxToXCell, 
            int catIdx,
            NumericValue cachedCatName,
            Category mainCategory) : this(catIdxToXCell, catIdx, cachedCatName)
        {
            // TODO: what about creating a new separate class like MultiCategory:Category
            MainCategory = mainCategory;
        }

        internal Category(
            ResettableLazy<List<X.Cell>> catIdxToXCell, 
            int catIdx,
            NumericValue cachedCatName)
        {
            _catIdxToXCell = catIdxToXCell;
            _catIdx = catIdx;
            _cachedCatName = cachedCatName;
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
        public string Name
        {
            get => _cachedCatName.InnerText;
            set
            {
                _catIdxToXCell.Value[_catIdx].CellValue.Text = value;
                var s =_catIdxToXCell.Value[_catIdx];
                //_cachedCatName = new NumericValue(value);
                _cachedCatName.Text = value;
                _catIdxToXCell.Reset();
            }
        }
#else
        /// <summary>
        ///     Gets category name.
        /// </summary>
        public string Name { get; }
#endif

        #endregion Properties
    }
}