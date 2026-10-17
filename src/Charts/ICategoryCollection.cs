using System.Collections.Generic;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of chart categories.
/// </summary>
public interface ICategoryCollection : IReadOnlyList<ICategory>
{
    /// <summary>
    ///     Appends a category to a non-combination chart with single-level text categories.
    /// </summary>
    /// <param name="name">The category name.</param>
    /// <param name="values">One value for each series, in series collection order.</param>
    /// <remarks>
    ///     Formula-backed data requires an embedded workbook and contiguous vertical ranges.
    /// </remarks>
    void Add(string name, params double[] values);
}