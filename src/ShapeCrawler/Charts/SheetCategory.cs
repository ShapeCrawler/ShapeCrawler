using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;

namespace ShapeCrawler.Charts;

internal sealed class SheetCategory : ICategory
{
    private readonly ChartPart sdkChartPart;
    private readonly string sheet;
    private readonly string address;
    private readonly NumericValue cachedValue;

    internal SheetCategory(ChartPart sdkChartPart, string sheet, string address, NumericValue cachedValue)
    {
        this.sdkChartPart = sdkChartPart;
        this.sheet = sheet;
        this.address = address;
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
            new ExcelBook(this.sdkChartPart).Sheet(this.sheet).UpdateCell(this.address, value);
        }
    }
}