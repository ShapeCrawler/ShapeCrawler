using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents table row.
/// </summary>
public interface ITableRow
{
    /// <summary>
    ///     Gets row's cells.
    /// </summary>
    IReadOnlyList<ITableCell> Cells { get; }

    /// <summary>
    ///     Gets or sets height in points.
    /// </summary>
    int Height { get; set; }

    /// <summary>
    ///     Creates a duplicate of the current row and adds this at the table end.
    /// </summary>
    ITableRow Clone();

    /// <summary>
    ///     Returns <see cref="A.TableRow" />.
    /// </summary>
    A.TableRow ATableRow();
}

internal sealed record SlideTableRow : ITableRow
{
    private readonly Lazy<List<TableCell>> cells;
    private readonly int index;

    internal SlideTableRow(SlidePart sdkSlidePart, A.TableRow aTableRow, int index)
    {
        this.ATableRow = aTableRow;
        this.index = index;
        this.cells = new Lazy<List<TableCell>>(() => this.GetCells());
    }

    public IReadOnlyList<ITableCell> Cells => this.cells.Value;

    public int Height
    {
        get => this.GetHeight();
        set => this.SetHeight(value);
    }

    internal A.TableRow ATableRow { get; }

    public ITableRow Clone()
    {
        var clonedRow = (A.TableRow)this.ATableRow.Clone();
        var addedRow = this.parentTable.AppendRow(clonedRow);

        return addedRow;
    }

    A.TableRow ITableRow.ATableRow()
    {
        return this.ATableRow;
    }

    private int GetHeight()
    {
        return (int)UnitConverter.EmuToPoint((int)this.ATableRow.Height!.Value);
    }
    
    private void SetHeight(int newPoints)
    {
        var currentPoints = this.GetHeight();
        if (currentPoints == newPoints)
        {
            return;
        }
        
        var newEmu = UnitConverter.PointToEmu(newPoints);
        this.ATableRow.Height!.Value = newEmu;

        if (newPoints > currentPoints)
        {
            var diffPoints = newPoints - currentPoints;
            var diffPixels = (int)UnitConverter.PointToPixel(diffPoints);
            this.parentTable.Height += diffPixels;
        }
        else
        {
            var diffPoints = currentPoints - newPoints;
            var diffPixels = (int)UnitConverter.PointToPixel(diffPoints);
            this.parentTable.Height -= diffPixels;
        }
    }
    
    private List<TableCell> GetCells()
    {
        var cellList = new List<TableCell?>();
        var aTcList = this.ATableRow.Elements<A.TableCell>();
        TableCell? addedCell = null;

        var columnIdx = 0;
        foreach (var aTc in aTcList)
        {
            if (aTc.HorizontalMerge is not null)
            {
                cellList.Add(addedCell);
            }
            else if (aTc.VerticalMerge is not null)
            {
                int upRowIdx = this.index - 1;
                TableCell upNeighborCell = (TableCell)this.parentTable[upRowIdx, columnIdx];
                cellList.Add(upNeighborCell);
                addedCell = upNeighborCell;
            }
            else
            {
                addedCell = new TableCell( aTc, this.index, columnIdx);
                cellList.Add(addedCell);
            }

            columnIdx++;
        }

        return cellList!;
    }
}