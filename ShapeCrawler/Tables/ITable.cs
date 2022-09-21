using System.Collections.Generic;
using ShapeCrawler.Collections;
using ShapeCrawler.Shapes;

// ReSharper disable CheckNamespace
namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a table on a slide.
    /// </summary>
    public interface ITable : IShape
    {
        /// <summary>
        ///     Gets table columns.
        /// </summary>
        IReadOnlyList<SCColumn> Columns { get; }

        /// <summary>
        ///     Gets table rows.
        /// </summary>
        RowCollection Rows { get; }

        /// <summary>
        ///     Gets cell by row and column indexes.
        /// </summary>
        ITableCell this[int rowIndex, int columnIndex] { get; }

        /// <summary>
        ///     Merge neighbor cells.
        /// </summary>
        void MergeCells(ITableCell inputCell1, ITableCell inputCell2);
    }
}