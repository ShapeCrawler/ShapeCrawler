using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Positions;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tables;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning disable IDE0130

/// <summary>
///     Represents a table row.
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
    decimal Height { get; set; }

    /// <summary>
    ///     Creates a duplicate of the current row and adds this at the table end.
    /// </summary>
    void Duplicate();
}

internal sealed class TableRow(A.TableRow aTableRow, int index): ITableRow
{
    public IReadOnlyList<ITableCell> Cells
    {
        get
        {
            var cells = new List<TableCell?>();
            var aTcList = this.ATableRow.Elements<A.TableCell>();
            TableCell? addedCell = null;

            var columnIdx = 0;
            foreach (var aTc in aTcList)
            {
                var mergedWithPreviousHorizontal = aTc.HorizontalMerge is not null;
                if (mergedWithPreviousHorizontal)
                {
                    cells.Add(addedCell);
                }
                else if (aTc.VerticalMerge is not null)
                {
                    var pGraphicFrame = this.ATableRow.Ancestors<P.GraphicFrame>().First();
                    var aTable = pGraphicFrame.GetFirstChild<A.Graphic>()!.GraphicData!.GetFirstChild<A.Table>()!;
                    var table = new Table(
                        new Shape(
                            new Position(pGraphicFrame),
                            new ShapeSize(pGraphicFrame),
                            new ShapeId(pGraphicFrame),
                            pGraphicFrame
                        ),
                        new TableRowCollection(pGraphicFrame),
                        new TableColumnCollection(pGraphicFrame),
                        new TableStyleOptions(aTable.TableProperties!),
                        pGraphicFrame
                    );
                    var upRowIdx = index - 1;
                    var upNeighborCell = (TableCell)table[upRowIdx, columnIdx];
                    cells.Add(upNeighborCell);
                    addedCell = upNeighborCell;
                }
                else
                {
                    addedCell = new TableCell(aTc, index, columnIdx);
                    cells.Add(addedCell);
                }

                columnIdx++;
            }

            return cells!;
        }
    }

    public decimal Height
    {
        get => new Emus(this.ATableRow.Height!.Value).AsPoints();
        set
        {
            var currentPoints = new Emus(this.ATableRow.Height!.Value).AsPoints();
            if (currentPoints == value)
            {
                return;
            }

            var newEmu = new Points(value).AsEmus();
            this.ATableRow.Height!.Value = newEmu;
            var pGraphicFrame = this.ATableRow.Ancestors<P.GraphicFrame>().First();
            var aTable = pGraphicFrame.GetFirstChild<A.Graphic>()!.GraphicData!.GetFirstChild<A.Table>()!;
            var parentTable = new Table(
                new Shape(
                    new Position(pGraphicFrame),
                    new ShapeSize(pGraphicFrame),
                    new ShapeId(pGraphicFrame),
                    pGraphicFrame
                ),
                new TableRowCollection(pGraphicFrame),
                new TableColumnCollection(pGraphicFrame),
                new TableStyleOptions(aTable.TableProperties!),
                pGraphicFrame
            );

            if (value > currentPoints)
            {
                var diffPoints = value - currentPoints;
                parentTable.SetTableHeight(parentTable.Height + diffPoints);
            }
            else
            {
                var diffPoints = currentPoints - value;
                parentTable.SetTableHeight(parentTable.Height - diffPoints);
            }
        }
    }

    internal A.TableRow ATableRow => aTableRow;

    public void Duplicate()
    {
        var rowCopy = (A.TableRow)this.ATableRow.Clone();
        this.ATableRow.Parent!.Append(rowCopy);
    }

    internal void SetHeight(int newPoints)
    {
        var currentPixels = new Emus(this.ATableRow.Height!.Value).AsPoints();

        if (currentPixels == newPoints)
        {
            return;
        }

        var newEmu = new Points(newPoints).AsEmus();
        this.ATableRow.Height!.Value = newEmu;
    }
}