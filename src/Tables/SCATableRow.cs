using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable InconsistentNaming
namespace ShapeCrawler.Tables;

internal readonly record struct SCATableRow
{
    private readonly A.TableRow aTableRow;

    internal SCATableRow(A.TableRow aTableRow)
    {
        this.aTableRow = aTableRow;
    }

    internal void AddNewCell()
    {
        var tableCell = CreateNewATableCell();

        this.aTableRow.Append(tableCell);
    }

    internal void InsertNewCellAfter(int columnNumber)
    {
        var cells = this.aTableRow.Elements<A.TableCell>().ToList();
        var targetCell = cells[columnNumber - 1];
        var aTableCell = CreateNewATableCell();

        this.aTableRow.InsertAfter(aTableCell, targetCell);
    }

    private static A.TableCell CreateNewATableCell()
    {
        var tableCell = new A.TableCell();
        var textBody = new A.TextBody();
        var bodyProperties = new A.BodyProperties();
        var listStyle = new A.ListStyle();
        var paragraph = new A.Paragraph();
        // Insert a default run so that paragraphs have at least one portion with font properties
        var runProperties = new A.RunProperties { Language = "en-US", FontSize = 1400, Dirty = false };
        var text = new A.Text { Text = string.Empty };
        var run = new A.Run(runProperties, text);
        var endParagraphRunProperties = new A.EndParagraphRunProperties { Language = "en-US" };

        paragraph.Append(run);
        paragraph.Append(endParagraphRunProperties);
        textBody.Append(bodyProperties);
        textBody.Append(listStyle);
        textBody.Append(paragraph);

        var tableCellProperties = new A.TableCellProperties();
        tableCell.Append(textBody);
        tableCell.Append(tableCellProperties);
        return tableCell;
    }
}