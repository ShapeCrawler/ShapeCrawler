using DocumentFormat.OpenXml.Packaging;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class Worksheet(EmbeddedPackagePart embeddedPackagePart, string sheetName)
{
    internal WorksheetCell Cell(string address) => new(embeddedPackagePart, sheetName, address);

    internal void UpdateCell(string address, string value) => this.Cell(address).UpdateValue(value);

    internal void UpdateCell(string address, string value, X.CellValues type) => this.Cell(address).UpdateValue(value, type);
}