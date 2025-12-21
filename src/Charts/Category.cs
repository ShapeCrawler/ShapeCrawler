using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class Category(
    ChartPart chartPart,
    NumericValue cachedValue,
    string? sheetName,
    string? cellAddress) : ICategory
{
    public bool HasMainCategory => false;
    
    public ICategory MainCategory => throw new SCException($"The main category is not available since the chart doesn't have a multi-category. " +
                                                           $"Use {nameof(ICategory.HasMainCategory)} property to check if the main category is available.");

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
