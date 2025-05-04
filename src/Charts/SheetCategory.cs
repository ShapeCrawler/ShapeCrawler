using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ShapeCrawler.Charts;

internal sealed class SheetCategory(
    ChartPart chartPart,
    string sheetName,
    string cellAddress,
    NumericValue cachedValue) : ICategory
{
    public bool HasMainCategory => false;

    public ICategory MainCategory => throw new SCException(
        $"The main category is not available since the chart doesn't have a multi-category. Use {nameof(ICategory.HasMainCategory)} property to check if the main category is available.");

    public string Name
    {
        get => cachedValue.InnerText;
        set
        {
            cachedValue.Text = value;
            new Workbook(chartPart.EmbeddedPackagePart!).Sheet(sheetName).UpdateCell(cellAddress, value, CellValues.String);
        }
    }
}