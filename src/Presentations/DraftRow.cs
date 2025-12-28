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

    /// <summary>
    ///     Gets or creates a cell at the specified index (1-based).
    /// </summary>
    public DraftCell Cell(int cellNumber)
    {
        if (cellNumber < 1)
        {
            throw new ArgumentException("Cell index must be 1-based and greater than 0.", nameof(cellNumber));
        }

        // Ensure we have enough cells
        while (this.cells.Count < cellNumber)
        {
            this.cells.Add(new DraftCell());
        }

        return this.cells[cellNumber - 1];
    }
}