using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Shared;
using ShapeCrawler.Wrappers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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
    void Duplicate();

    /// <summary>
    ///     Returns <see cref="A.TableRow" />.
    /// </summary>
    A.TableRow ATableRow();
}

internal sealed class SlideTableRow : ITableRow
{
    private readonly Lazy<List<TableCell>> cells;
    private readonly SlidePart sdkSlidePart;
    private readonly int index;

    internal SlideTableRow(SlidePart sdkSlidePart, A.TableRow aTableRow, int index)
    {
        this.ATableRow = aTableRow;
        this.sdkSlidePart = sdkSlidePart;
        this.index = index;
        this.cells = new Lazy<List<TableCell>>(() => this.GetCells());
    }

    public IReadOnlyList<ITableCell> Cells => this.cells.Value;

    public int Height
    {
        get => this.GetHeight();
        set => this.UpdateHeight(value);
    }

    internal A.TableRow ATableRow { get; }

    public void Duplicate()
    {
        var rowCopy = (A.TableRow)this.ATableRow.Clone();
        this.ATableRow.Parent!.Append(rowCopy);
    }

    A.TableRow ITableRow.ATableRow()
    {
        return this.ATableRow;
    }

    private int GetHeight()
    {
        return (int)UnitConverter.EmuToPoint((int)this.ATableRow.Height!.Value);
    }
    
    private void UpdateHeight(int newPoints)
    {
        var currentPoints = this.GetHeight();
        if (currentPoints == newPoints)
        {
            return;
        }
        
        var newEmu = UnitConverter.PointToEmu(newPoints);
        this.ATableRow.Height!.Value = newEmu;

        var pGraphicalFrame = new SdkOpenXmlElement(ATableRow).FirstAncestor<P.GraphicFrame>();
        var parentTable = new SlideTable(this.sdkSlidePart, pGraphicalFrame);
        if (newPoints > currentPoints)
        {
            var diffPoints = newPoints - currentPoints;
            var diffPixels = (int)UnitConverter.PointToPixel(diffPoints);
            parentTable.Height += diffPixels;
        }
        else
        {
            var diffPoints = currentPoints - newPoints;
            var diffPixels = (int)UnitConverter.PointToPixel(diffPoints);
            parentTable.Height -= diffPixels;
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
                var pGraphicalFrame = new SdkOpenXmlElement(ATableRow).FirstAncestor<P.GraphicFrame>();
                var parentTable = new SlideTable(this.sdkSlidePart, pGraphicalFrame);
                int upRowIdx = this.index - 1;
                var upNeighborCell = (TableCell)parentTable[upRowIdx, columnIdx];
                cellList.Add(upNeighborCell);
                addedCell = upNeighborCell;
            }
            else
            {
                addedCell = new TableCell(this.sdkSlidePart,aTc, this.index, columnIdx);
                cellList.Add(addedCell);
            }

            columnIdx++;
        }

        return cellList!;
    }
}