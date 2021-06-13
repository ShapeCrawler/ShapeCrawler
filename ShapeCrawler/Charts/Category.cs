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
        private readonly int index;
        private readonly NumericValue cachedName;
        private readonly ResettableLazy<List<X.Cell>> indexToXCell;

        #region Constructors

        internal Category(
            ResettableLazy<List<X.Cell>> indexToXCell,
            int index,
            NumericValue cachedName,
            Category mainCategory)
            : this(indexToXCell, index, cachedName)
        {
            // TODO: what about creating a new separate class like MultiCategory:Category
            this.MainCategory = mainCategory;
        }

        internal Category(
            ResettableLazy<List<X.Cell>> indexToXCell,
            int index,
            NumericValue cachedName)
        {
            this.indexToXCell = indexToXCell;
            this.index = index;
            this.cachedName = cachedName;
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
            get => this.cachedName.InnerText;
            set
            {
                this.indexToXCell.Value[index].CellValue.Text = value;
                this.cachedName.Text = value;
                this.indexToXCell.Reset();
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