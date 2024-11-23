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
            var aTableCell = new SATableCell().ATableCell();
            row.Append(aTableCell);
        }
        
        aTable.Append(row);

        return row;
    }
}