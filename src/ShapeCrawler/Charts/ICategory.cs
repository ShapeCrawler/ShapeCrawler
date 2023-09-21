using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Charts;
using ShapeCrawler.Shared;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a chart category.
/// </summary>
public interface ICategory
{
    /// <summary>
    ///     Gets main category. Returns <c>NULL</c> if chart is not Multi-Category.
    /// </summary>
    public ICategory? MainCategory { get; }

    /// <summary>
    ///     Gets or sets category name.
    /// </summary>
    string Name { get; set; }
}

internal sealed class Category : ICategory
{
    private readonly ChartPart sdkChartPart;
    private readonly Formula cFormula;
    private readonly int index;
    private readonly NumericValue cachedValue;

    internal Category(ChartPart sdkChartPart, int index, NumericValue cachedValue, Category mainCategory)
    {
        this.index = index;
        this.cachedValue = cachedValue;
        this.MainCategory = mainCategory;
    }
    
    internal Category(ChartPart sdkChartPart, int index, NumericValue cachedValue)
    {
        this.index = index;
        this.cachedValue = cachedValue;
    }

    internal Category(ChartPart sdkChartPart, C.Formula cFormula, int index, NumericValue cachedValue)
    {
        this.sdkChartPart = sdkChartPart;
        this.cFormula = cFormula;
        this.index = index;
        this.cachedValue = cachedValue;
    }

    public ICategory MainCategory { get; }
    
    public string Name
    {
        get => this.cachedValue.InnerText;
        set
        {
            if (this.MainCategory != null)
            {
                const string msg =
                    "Sorry, but updating the category name of Multi-Category charts have not yet been supported by ShapeCrawler." +
                    "If it is critical for you, you are always welcome for this implementation. " +
                    "We will wait for your Pull Request on https://github.com/ShapeCrawler/ShapeCrawler.";
                throw new NotSupportedException(msg);
            }

            this.cachedValue.Text = value;

            var xCells = this.FormulaCells();
            var xCell = xCells[this.index];
            xCell.DataType = new DocumentFormat.OpenXml.EnumValue<X.CellValues>(X.CellValues.String);
            xCell.CellValue = new X.CellValue(value);
        }
    }

    private List<X.Cell> FormulaCells()
    {
        var normalizedFormula = cFormula.Text.Replace("'", string.Empty).Replace("$", string.Empty); // eg: Sheet1!$A$2:$A$5 -> Sheet1!A2:A5
        var chartSheetName = Regex.Match(normalizedFormula, @".+(?=\!)").Value; // eg: Sheet1!A2:A5 -> Sheet1
        var cellsRange = Regex.Match(normalizedFormula, @"(?<=\!).+").Value; // eg: Sheet1!A2:A5 -> A2:A5

        var workbookPart = slideChart.workbook!.WorkbookPart;
        var xSheet = workbookPart.Workbook.Sheets!.Elements<X.Sheet>().First(xSheet => xSheet.Name == chartSheetName);
        var sdkWorksheetPart = (WorksheetPart)workbookPart.GetPartById(xSheet.Id!);
        var xCells = sdkWorksheetPart.Worksheet.Descendants<X.Cell>();

        var rangeCellAddresses = new CellsRangeParser(cellsRange).GetCellAddresses();
        var rangeXCells = new List<X.Cell>(rangeCellAddresses.Count);
        foreach (var address in rangeCellAddresses)
        {
            var xCell = xCells.First(xCell => xCell.CellReference == address);
            rangeXCells.Add(xCell);
        }

        return rangeXCells;
    }
}