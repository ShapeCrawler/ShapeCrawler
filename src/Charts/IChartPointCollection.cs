using System.Collections.Generic;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a collection of numeric chart points.
/// </summary>
public interface IChartPointCollection : IReadOnlyList<IChartPoint>
{
    /// <summary>
    ///     Appends a numeric value to this collection.
    /// </summary>
    /// <param name="value">The point value.</param>
    /// <remarks>
    ///     Formula-backed data requires an embedded workbook, a contiguous vertical range, and an unoccupied next cell.
    ///     This method extends only this collection; add matching X values and bubble sizes separately.
    ///     To append a category and values for all series, use <see cref="ICategoryCollection.Add"/>.
    /// </remarks>
    void Add(double value);
}