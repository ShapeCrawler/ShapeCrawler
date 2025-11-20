using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ShapeCrawler.Extensions;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a table row collection.
/// </summary>
public interface ITableRowCollection : IEnumerable<ITableRow>
{
    /// <summary>
    ///     Gets number of rows.
    /// </summary>
    int Count { get; }

    /// <summary>
    ///     Gets row by index.
    /// </summary>
    ITableRow this[int index] { get; }

    /// <summary>
    ///     Removes specified row from collection.
    /// </summary>
    void Remove(ITableRow removing);

    /// <summary>
    ///     Removes table row by index.
    /// </summary>
    void RemoveAt(int index);

    /// <summary>
    ///     Adds a new row at the end of the table.
    /// </summary>
    void Add();
    
    /// <summary>
    ///     Adds a new row at the specified index.
    /// </summary>
    void Add(int index);

    /// <summary>
    ///     Adds a new row at the specified index.
    /// </summary>
    /// <param name="index">Index where the new row will be added.</param>
    /// <param name="templateRowIndex">Index of the row used as a template.</param>
    void Add(int index, int templateRowIndex);
}

internal sealed class TableRowCollection : ITableRowCollection
{
    private readonly A.Table aTable;

    internal TableRowCollection(P.GraphicFrame pGraphicFrame)
    {
        this.aTable = pGraphicFrame.GetFirstChild<A.Graphic>() !.GraphicData!.GetFirstChild<A.Table>() !;
    }

    public int Count => this.Rows().Count;

    public ITableRow this[int index] => this.Rows()[index];

    public void Remove(ITableRow removing)
    {
        var removingRowInternal = (TableRow)removing;
        removingRowInternal.ATableRow.Remove();
    }

    public void RemoveAt(int index)
    {
        var rows = this.Rows();
        if (index < 0 || index >= rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var innerRow = rows[index];
        this.Remove(innerRow);
    }

    public void Add()
    {
        var columnsCount = this.Rows()[0].Cells.Count;
        var aTableRow = new A.TableRow { Height = Constants.DefaultRowHeightEmu };
        for (var i = 0; i < columnsCount; i++)
        {
            new SCATableRow(aTableRow).AddNewCell();
        }
        
        aTable.Append(aTableRow);
    }

    public void Add(int index)
    {
        var rows = this.Rows();
        if (index < 0 || index > rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        var columnsCount = rows.Count > 0 ? rows[0].Cells.Count : 0;
        if (columnsCount == 0)
        {
            throw new InvalidOperationException("Cannot add a row to an empty table.");
        }

        var aTableRow = new A.TableRow { Height = Constants.DefaultRowHeightEmu };
        for (var i = 0; i < columnsCount; i++)
        {
            new SCATableRow(aTableRow).AddNewCell();
        }
        
        // Get the element before which we want to insert the new row
        var aTableRows = this.aTable.Elements<A.TableRow>().ToList();
        if (index == aTableRows.Count)
        {
            // Add at the end
            this.aTable.Append(aTableRow);
        }
        else
        {
            // Insert before the row at the specified index
            this.aTable.InsertBefore(aTableRow, aTableRows[index]);
        }
    }

    IEnumerator<ITableRow> IEnumerable<ITableRow>.GetEnumerator() => this.Rows().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => this.Rows().GetEnumerator();
    
    public void Add(int index, int templateRowIndex)
    {
        var rows = this.Rows();
        if (index < 0 || index > rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        if (templateRowIndex < 0 || templateRowIndex >= rows.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(templateRowIndex));
        }

        // Get template row properties
        var templateRow = rows[templateRowIndex];
        var templateARow = templateRow.ATableRow;

        // Create a new row with the same height as the template
        var newARow = new A.TableRow { Height = templateARow.Height };
        
        var templateACells = templateARow.Elements<A.TableCell>().ToList();
        
        // Build each cell of the new row based on the template cell formatting
        foreach (var (templateACell, columnIndex) in templateACells.Select((c, i) => (c, i)))
        {
            var templateCell = templateRow.Cells[columnIndex];
            var newACell = CreateCellFromTemplate(templateACell, templateCell);
            newARow.Append(newACell);
        }
        
        // Get the element before which we want to insert the new row
        var aTableRows = this.aTable.Elements<A.TableRow>().ToList();
        if (index == aTableRows.Count)
        {   
            // Add at the end
            this.aTable.Append(newARow);
        }
        else
        {   
            // Insert before the row at the specified index
            this.aTable.InsertBefore(newARow, aTableRows[index]);
        }
    }

    private static A.TableCell CreateCellFromTemplate(A.TableCell templateACell, ITableCell templateCell)
    {
        // Create a brand-new table cell with an empty text body
        var newACell = new A.TableCell();
        var textBody = new A.TextBody(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" }));
        newACell.Append(textBody);
        
        // Determine template cell properties and font color
        var templateFontColor = templateCell.TextBox.Paragraphs.FirstOrDefault()?.Portions.FirstOrDefault()?.Font?.Color.Hex;
        
        A.TableCellProperties newTcPr;
        if (templateACell.TableCellProperties is not null)
        {
            // Clone existing TableCellProperties from template
            newTcPr = (A.TableCellProperties)templateACell.TableCellProperties.CloneNode(true);
        }
        else
        {
            newTcPr = new A.TableCellProperties();
        }
        
        if (!string.IsNullOrEmpty(templateFontColor))
        {
            newTcPr.AddSolidFill(templateFontColor!);
        }
        else
        {
            newTcPr.AddSolidFill("000000"); // default color
        }

        newACell.Append(newTcPr);
        
        return newACell;
    }

    private List<TableRow> Rows() =>
    [
        .. this.aTable.Elements<A.TableRow>().Select((aTableRow, index) => new TableRow(aTableRow, index))
    ];
}