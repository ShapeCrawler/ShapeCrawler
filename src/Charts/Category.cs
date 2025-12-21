using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class Category : ICategory
{
    private readonly NumericValue cachedValue;
    private readonly ChartPart chartPart;
    private readonly string? sheetName;
    private readonly string? cellAddress;

    internal Category(
        ChartPart chartPart,
        NumericValue cachedValue,
        string? sheetName,
        string? cellAddress)
    {
        this.chartPart = chartPart;
        this.cachedValue = cachedValue;
        this.sheetName = sheetName;
        this.cellAddress = cellAddress;
    }

    public bool HasMainCategory => false;
    
    public ICategory MainCategory => throw new SCException($"The main category is not available since the chart doesn't have a multi-category. " +
                                                           $"Use {nameof(ICategory.HasMainCategory)} property to check if the main category is available.");

    public string Name
    {
        get => this.cachedValue.InnerText;
        set
        {
            this.cachedValue.Text = value;
            if (this.sheetName != null && 
                this.cellAddress != null && 
                this.chartPart.EmbeddedPackagePart != null)
            {
                new Workbook(this.chartPart.EmbeddedPackagePart).Sheet(this.sheetName).UpdateCell(this.cellAddress, value, CellValues.String);
            }
        }
    }
}
