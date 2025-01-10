using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

/// <summary>
///     Represents a table style options.
/// </summary>
public interface ITableStyleOptions
{
    /// <summary>
    ///     Gets or sets a value indicating whether the table has header row.
    /// </summary>
    public bool HasHeaderRow { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether the table has total row.
    /// </summary>
    public bool HasTotalRow { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether the table has banded rows.
    /// </summary>
    public bool HasBandedRows { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether the table has first column.
    /// </summary>
    public bool HasFirstColumn { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether the table has last column.
    /// </summary>
    public bool HasLastColumn { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether the table has banded columns.
    /// </summary>
    public bool HasBandedColumns { get; set; }
}

internal sealed class TableStyleOptions(TableProperties tableProperties)
    : ITableStyleOptions
{
    public bool HasHeaderRow
    {
        get => tableProperties.FirstRow?.Value ?? false;
        set => tableProperties.FirstRow = value;
    }

    public bool HasTotalRow
    {
        get => tableProperties.LastRow?.Value ?? false;
        set => tableProperties.LastRow = value;
    }

    public bool HasBandedRows
    {
        get => tableProperties.BandRow?.Value ?? false;
        set => tableProperties.BandRow = value;
    }

    public bool HasFirstColumn
    {
        get => tableProperties.FirstColumn?.Value ?? false;
        set => tableProperties.FirstColumn = value;
    }

    public bool HasLastColumn
    {
        get => tableProperties.LastColumn?.Value ?? false;
        set => tableProperties.LastColumn = value;
    }

    public bool HasBandedColumns
    {
        get => tableProperties.BandColumn?.Value ?? false;
        set => tableProperties.BandColumn = value;
    }
}