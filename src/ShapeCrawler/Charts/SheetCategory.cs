using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ShapeCrawler.Excel;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Charts;

internal sealed class SheetCategory : ICategory
{
    private readonly ChartPart sdkChartPart;
    private readonly string sheetName;
    private readonly string cellAddress;
    private readonly NumericValue cachedValue;

    internal SheetCategory(ChartPart sdkChartPart, string sheetName, string cellAddress, NumericValue cachedValue)
    {
        this.sdkChartPart = sdkChartPart;
        this.sheetName = sheetName;
        this.cellAddress = cellAddress;
        this.cachedValue = cachedValue;
    }

    public bool HasMainCategory => false;
    
    public ICategory MainCategory => throw new SCException($"The main category is not available since the chart doesn't have a multi-category. Use {nameof(ICategory.HasMainCategory)} property to check if the main category is available.");

    public string Name
    {
        get => this.cachedValue.InnerText;
        set
        {
            this.cachedValue.Text = value;
            new ExcelBook(this.sdkChartPart).Sheet(this.sheetName).UpdateCell(this.cellAddress, value, CellValues.String);
        }
    }
}