using System;
using System.Collections.Generic;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft table builder.
/// </summary>
public sealed class DraftTable
{
    private readonly List<DraftRow> rows = [];

    internal int ColumnsCount { get; private set; } = 2;

    internal int TableX { get; private set; } = 0;

    internal int TableY { get; private set; } = 0;

    internal IReadOnlyList<DraftRow> Rows => this.rows;

    /// <summary>
    ///     Sets the number of columns in the table.
    /// </summary>
    public DraftTable Columns(int count)
    {
        this.ColumnsCount = count;
        return this;
    }

    /// <summary>
    ///     Adds a row to the table with configuration.
    /// </summary>
    public DraftTable Row(Action<DraftRow> configure)
    {
        var rowBuilder = new DraftRow();
        configure(rowBuilder);
        this.rows.Add(rowBuilder);
        return this;
    }
}
