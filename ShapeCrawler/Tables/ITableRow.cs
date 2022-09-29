using System;
using System.Collections.Generic;
using ShapeCrawler.Exceptions;
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
    ///     Creates a duplicate of the current row.
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

#if DEBUG
    public ITableRow Clone()
    {
        var clonedRow = (A.TableRow)this.ATableRow.Clone();
        var addedRow = this.ParentTable.AppendRow(clonedRow);

        return addedRow;
    }
#endif

    internal void ThrowIfRemoved()
    {
        
    }
    
    #region Private Methods

    private List<SCTableCell> GetCells()
    {
        var cellList = new List<SCTableCell>();
        IEnumerable<A.TableCell> aTableCells = this.ATableRow.Elements<A.TableCell>();
        SCTableCell addedScCell = null;

        int columnIdx = 0;
        foreach (A.TableCell aTableCell in aTableCells)
        {
            if (aTableCell.HorizontalMerge != null)
            {
                cellList.Add(addedScCell);
            }
            else if (aTableCell.VerticalMerge != null)
            {
                int upRowIdx = this.index - 1;
                SCTableCell upNeighborScCell = (SCTableCell)this.ParentTable[upRowIdx, columnIdx];
                cellList.Add(upNeighborScCell);
                addedScCell = upNeighborScCell;
            }
            else
            {
                addedScCell = new SCTableCell(this, aTableCell, this.index, columnIdx);
                cellList.Add(addedScCell);
            }

            columnIdx++;
        }

        return cellList;
    }

    #endregion
}