using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
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

        /// <summary>
        /// Adds a row to the bottom of the table.
        /// </summary>
        /// <param name="row">Row that will be added to the table.</param>
        /// <returns>A reference to the recently added row.</returns>
        SCTableRow AppendRow(TableRow row);

        void RemoveRowAt(int index);


#if DEBUG
        public object GetInternalObject();

#endif
    }
}