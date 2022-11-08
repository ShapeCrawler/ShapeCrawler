using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing.Charts;
using ShapeCrawler.Shared;
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

internal class Category : ICategory
{
    private readonly int index;
    private readonly NumericValue cachedValue;
    private readonly ResettableLazy<List<X.Cell>> xCells;

    internal Category(
        ResettableLazy<List<X.Cell>> xCells,
        int index,
        NumericValue cachedValue,
        Category mainCategory)
        : this(xCells, index, cachedValue)
    {
        // TODO: what about creating a new separate class like MultiCategory:Category
        this.MainCategory = mainCategory;
    }

    internal Category(ResettableLazy<List<X.Cell>> xCells, int index, NumericValue cachedValue)
    {
        this.xCells = xCells;
        this.index = index;
        this.cachedValue = cachedValue;
    }

    public ICategory? MainCategory { get; }

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

            var xCell = this.xCells.Value[this.index];
            xCell.DataType = new DocumentFormat.OpenXml.EnumValue<X.CellValues>(X.CellValues.String);
            xCell.CellValue = new X.CellValue(value);

            this.xCells.Reset();
        }
    }
}