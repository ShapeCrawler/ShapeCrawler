using System;
using System.Collections.Generic;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a row in a table.
/// </summary>
public interface ITableRow
{
    /// <summary>
    ///     Gets row's cells.
    /// </summary>
    IReadOnlyList<ITableCell> Cells { get; }

    /// <summary>
    ///     Gets or sets height.
    /// </summary>
    long Height { get; set; }

    /// <summary>
    ///     Creates a duplicate of the current row and adds this at the table end.
    /// </summary>
    ITableRow Clone();
}

internal class SCTableRow : ITableRow
{
    private readonly Lazy<List<SCTableCell>> cells;
    private readonly int index;

    internal SCTableRow(SlideTable table, A.TableRow aTableRow, int index)
    {
        this.ParentTable = table;
        this.ATableRow = aTableRow;
        this.index = index;

#if NETSTANDARD2_0
        this.cells = new Lazy<List<SCTableCell>>(() => this.GetCells());
#else
        this.cells = new Lazy<List<SCTableCell>>(this.GetCells);
#endif
    }

    public IReadOnlyList<ITableCell> Cells => this.cells.Value;

    public long Height
    {
        get => this.ATableRow.Height!.Value;
        set => this.ATableRow.Height!.Value = value;
    }

    internal SlideTable ParentTable { get; }

    internal A.TableRow ATableRow { get; }

    public ITableRow Clone()
    {
        var clonedRow = (A.TableRow)this.ATableRow.Clone();
        var addedRow = this.ParentTable.AppendRow(clonedRow);

        return addedRow;
    }

    internal void ThrowIfRemoved()
    {

    }

    #region Private Methods

    private List<SCTableCell> GetCells()
    {
        var cellList = new List<SCTableCell?>();
        var aCells = this.ATableRow.Elements<A.TableCell>();
        SCTableCell? addedCell = null;

        var columnIdx = 0;
        foreach (var aTableCell in aCells)
        {
            if (aTableCell.HorizontalMerge is not null)
            {
                cellList.Add(addedCell);
            }
            else if (aTableCell.VerticalMerge is not null)
            {
                int upRowIdx = this.index - 1;
                SCTableCell upNeighborScCell = (SCTableCell)this.ParentTable[upRowIdx, columnIdx];
                cellList.Add(upNeighborScCell);
                addedCell = upNeighborScCell;
            }
            else
            {
                addedCell = new SCTableCell(this, aTableCell, this.index, columnIdx);
                cellList.Add(addedCell);
            }

            columnIdx++;
        }

        return cellList;
    }

    #endregion
}