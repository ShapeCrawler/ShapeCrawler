using System.Collections.Generic;
using ShapeCrawler.Collections;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;

namespace ShapeCrawler.Tables
{
    public interface ITable : IShape
    {
        IReadOnlyList<Column> Columns { get; }
        RowCollection Rows { get; }
        CellSc this[int rowIndex, int columnIndex] { get; }
    }
}