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
        private ResettableLazy<List<Cell>> xCells;
        private int xCellIdx;
        private NumericValue cachedValue;
        private string catName;
        private Category value;
        private ResettableLazy<Dictionary<int, Cell>> catIndexToXCell;
        private int v;


        #region Constructors

        public Category(string catName)
        {
            this.catName = catName;
        }

        public Category(string catName, Category value)
        {
            this.catName = catName;
            this.value = value;
        }

        public Category(ResettableLazy<List<Cell>> xCells, int xCellIdx, NumericValue cachedValue)
        {
            this.xCells = xCells;
            this.xCellIdx = xCellIdx;
            this.cachedValue = cachedValue;
        }

        public Category(ResettableLazy<Dictionary<int, Cell>> catIndexToXCell, int v, NumericValue cachedValue)
        {
            this.catIndexToXCell = catIndexToXCell;
            this.v = v;
            this.cachedValue = cachedValue;
        }

        /// <summary>
        ///     Initializes non-multi-category.
        /// </summary>
        internal Category(Shared.ResettableLazy<System.Collections.Generic.List<DocumentFormat.OpenXml.Spreadsheet.Cell>> xCells, int xCellIdx, string name)
        {
            Name = name;
        }

        /// <summary>
        ///     Initializes multi-category.
        /// </summary>
        internal Category(Shared.ResettableLazy<System.Collections.Generic.List<DocumentFormat.OpenXml.Spreadsheet.Cell>> xCells, string name, Category main)
        {
            Name = name;
            MainCategory = main;
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