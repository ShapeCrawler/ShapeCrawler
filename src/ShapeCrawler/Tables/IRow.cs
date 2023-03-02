using System;
using System.Collections.Generic;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents table row.
/// </summary>
public interface IRow
{
    /// <summary>
    ///     Gets row's cells.
    /// </summary>
    IReadOnlyList<ICell> Cells { get; }

    /// <summary>
    ///     Gets or sets height.
    /// </summary>
    long Height { get; set; }

    /// <summary>
    ///     Creates a duplicate of the current row and adds this at the table end.
    /// </summary>
    IRow Clone();
}

internal sealed class SCRow : IRow
{
    private readonly Lazy<List<SCCell>> cells;
    private readonly int index;

    internal SCRow(SCTable table, A.TableRow aTableRow, int index)
    {
        this.ParentTable = table;
        this.ATableRow = aTableRow;
        this.index = index;
        this.cells = new Lazy<List<SCCell>>(() => this.GetCells());
    }

    public IReadOnlyList<ICell> Cells => this.cells.Value;

    public long Height
    {
        get => this.ATableRow.Height!.Value;
        set => this.ATableRow.Height!.Value = value;
    }

    internal SCTable ParentTable { get; }

    internal A.TableRow ATableRow { get; }

    public IRow Clone()
    {
        var clonedRow = (A.TableRow)this.ATableRow.Clone();
        var addedRow = this.ParentTable.AppendRow(clonedRow);

        return addedRow;
    }

    private List<SCCell> GetCells()
    {
        var cellList = new List<SCCell?>();
        var aCells = this.ATableRow.Elements<A.TableCell>();
        SCCell? addedCell = null;

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
                SCCell upNeighborScCell = (SCCell)this.ParentTable[upRowIdx, columnIdx];
                cellList.Add(upNeighborScCell);
                addedCell = upNeighborScCell;
            }
            else
            {
                addedCell = new SCCell(this, aTableCell, this.index, columnIdx);
                cellList.Add(addedCell);
            }

            columnIdx++;
        }

        return cellList!;
    }
}