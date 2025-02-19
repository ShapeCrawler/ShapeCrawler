using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions;

internal static class ATableExtensions
{
    internal static A.TableRow AddRow(this A.Table aTable, int columnsCount)
    {
        var aTableRow = new A.TableRow { Height = Constants.DefaultRowHeightEmu };
        for (var i = 0; i < columnsCount; i++)
        {
            new SCATableRow(aTableRow).AddNewCell();
        }
        
        aTable.Append(aTableRow);

        return aTableRow;
    }
}