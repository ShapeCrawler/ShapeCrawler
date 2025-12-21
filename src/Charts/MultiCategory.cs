using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class MultiCategory(
    ChartPart chartPart,
    ICategory mainCategory,
    NumericValue cachedValue,
    string? sheetName,
    string? cellAddress) : ICategory
{
    public bool HasMainCategory => true;
    
    public ICategory MainCategory { get; } = mainCategory;

    public string Name
    {
        get => cachedValue.InnerText;
        set
        {
            cachedValue.Text = value;
            if (sheetName != null && 
                cellAddress != null && 
                chartPart.EmbeddedPackagePart != null)
            {
                new Workbook(chartPart.EmbeddedPackagePart).Sheet(sheetName).UpdateCell(cellAddress, value, CellValues.String);
            }
        }
    }
}
