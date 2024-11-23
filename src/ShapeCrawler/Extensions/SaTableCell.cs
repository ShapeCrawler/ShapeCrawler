using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal record struct SaTableCell
{
    internal static A.TableCell ATableCell()
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

        return tableCell;
    }
}