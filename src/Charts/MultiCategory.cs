using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class MultiCategory : ICategory
{
    private readonly NumericValue cachedValue;
    private readonly ChartPart chartPart;
    private readonly string? sheetName;
    private readonly string? cellAddress;

    internal MultiCategory(
        ChartPart chartPart, 
        ICategory mainCategory, 
        NumericValue cachedValue,
        string? sheetName,
        string? cellAddress)
    {
        this.chartPart = chartPart;
        this.MainCategory = mainCategory;
        this.cachedValue = cachedValue;
        this.sheetName = sheetName;
        this.cellAddress = cellAddress;
    }

    public bool HasMainCategory => true;
    
    public ICategory MainCategory { get; }

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
