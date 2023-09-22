using System;
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
    ///     Gets a value indicating whether the category has a main category.
    /// </summary>
    public bool HasMainCategory { get; }
 
    /// <summary>
    ///     Returns main category.
    /// </summary>
    public ICategory MainCategory { get; }

    /// <summary>
    ///     Gets or sets category name.
    /// </summary>
    string Name { get; set; }
}