using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class ExcelCells
{
    private readonly ChartPart sdkChartPart;
    private readonly C.Formula cFormula;

    internal ExcelCells(ChartPart sdkChartPart, C.Formula cFormula)
    {
        this.sdkChartPart = sdkChartPart;
        this.cFormula = cFormula;
    }

    internal string ContentCellAt(int index)
    {
        var normalizedFormula = cFormula.Text.Replace("'", string.Empty).Replace("$", string.Empty); // eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5
        var chartSheetName = Regex.Match(normalizedFormula, @".+(?=\!)").Value; // eg: Sheet1!A2:A5 -> Sheet1
        var cellsRange = Regex.Match(normalizedFormula, @"(?<=\!).+").Value; // eg: Sheet1!A2:A5 -> A2:A5

        var stream = this.sdkChartPart.EmbeddedPackagePart!.GetStream();
        var sdkSpreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var workbookPart = sdkSpreadsheetDocument.WorkbookPart!;
        var chartXSheet = workbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == chartSheetName);
        var sdkWorksheetPart = (WorksheetPart)workbookPart.GetPartById(chartXSheet.Id!);
        var xCells = sdkWorksheetPart.Worksheet.Descendants<X.Cell>();

        var rangeCellAddresses = new CellsRangeParser(cellsRange).GetCellAddresses();
        var rangeXCells = new List<X.Cell>(rangeCellAddresses.Count);
        foreach (var address in rangeCellAddresses)
        {
            var xCell = xCells.First(xCell => xCell.CellReference == address);
            rangeXCells.Add(xCell);
        }

        return rangeXCells[index].InnerText;
    }
}