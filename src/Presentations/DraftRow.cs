using System;
using System.Collections.Generic;

namespace ShapeCrawler.Presentations;

/// <summary>
///     Represents a draft row builder.
/// </summary>
public sealed class DraftRow
{
    private readonly List<DraftCell> cells = [];

    internal IReadOnlyList<DraftCell> Cells => this.cells;

    /// <summary>
    ///     Adds a cell to the row.
    /// </summary>
    public DraftRow Cell()
    {
        this.cells.Add(new DraftCell());
        return this;
    }

    /// <summary>
    ///     Adds a cell to the row with configuration.
    /// </summary>
    public DraftRow Cell(Action<DraftCell> configure)
    {
        var cellBuilder = new DraftCell();
        configure(cellBuilder);
        this.cells.Add(cellBuilder);
        return this;
    }
}