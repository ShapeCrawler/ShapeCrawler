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
        private readonly int _index;
        private readonly NumericValue _cachedName;
        private readonly ResettableLazy<List<X.Cell>> _indexToXCell;

        #region Constructors

        internal Category(
            ResettableLazy<List<X.Cell>> indexToXCell,
            int index,
            NumericValue cachedName,
            Category mainCategory) : this(indexToXCell, index, cachedName)
        {
            // TODO: what about creating a new separate class like MultiCategory:Category
            MainCategory = mainCategory;
        }

        internal Category(
            ResettableLazy<List<X.Cell>> indexToXCell,
            int index,
            NumericValue cachedName)
        {
            _indexToXCell = indexToXCell;
            _index = index;
            _cachedName = cachedName;
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
            get => this._cachedName.InnerText;
            set
            {
                _indexToXCell.Value[_index].CellValue.Text = value;
                _cachedName.Text = value;
                _indexToXCell.Reset();
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