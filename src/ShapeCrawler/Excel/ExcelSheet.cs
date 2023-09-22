using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Excel;

internal record struct ExcelSheet
{
    private readonly ChartPart sdkChartPart;
    private readonly string sheetName;

    internal ExcelSheet(ChartPart sdkChartPart, string sheetName)
    {
        this.sdkChartPart = sdkChartPart;
        this.sheetName = sheetName;
    }

    internal void UpdateCell(string address, string value)
    {
        throw new System.NotImplementedException();
    }
}