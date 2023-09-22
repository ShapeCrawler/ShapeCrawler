using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class ATableExtensions
{
    internal static A.TableRow AddRow(this A.Table aTable, int columnsCount)
    {
        var row = new A.TableRow { Height = Constants.DefaultRowHeightEmu };
        for (var i = 0; i < columnsCount; i++)
        {
            var tableCell = new A.TableCell();
            var textBody = new A.TextBody();
            var bodyProperties = new A.BodyProperties();
            var listStyle = new A.ListStyle();
            var paragraph = new A.Paragraph();
            var endParagraphRunProperties = new A.EndParagraphRunProperties { Language = "en-US" };
            paragraph.Append(endParagraphRunProperties);
            textBody.Append(bodyProperties);
            textBody.Append(listStyle);
            textBody.Append(paragraph);
            var tableCellProperties = new A.TableCellProperties();
            tableCell.Append(textBody);
            tableCell.Append(tableCellProperties);

            row.Append(tableCell);
        }
        
        aTable.Append(row);

        return row;
    }
}