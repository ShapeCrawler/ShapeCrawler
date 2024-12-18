using DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tables;

/// <summary>
///     Represents table style options.
/// </summary>
public interface ITableStyleOptions
{
    /// <summary>
    ///     Gets or sets a value indicating whether table has header row.
    /// </summary>
    public bool HasHeaderRow { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether table has total row.
    /// </summary>
    public bool HasTotalRow { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether table has banded rows.
    /// </summary>
    public bool HasBandedRows { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether table has first column.
    /// </summary>
    public bool HasFirstColumn { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether table has last column.
    /// </summary>
    public bool HasLastColumn { get; set; }
    
    /// <summary>
    ///     Gets or sets a value indicating whether table has banded columns.
    /// </summary>
    public bool HasBandedColumns { get; set; }
}

/// <summary>
///    Represents table style options.
/// </summary>
internal sealed class TableStyleOptions(TableProperties tableProperties)
    : ITableStyleOptions
{
    /// <inheritdoc/>
    public bool HasHeaderRow
    {
        get => tableProperties.FirstRow?.Value ?? false;
        set => tableProperties.FirstRow = value;
    }

    /// <inheritdoc/>
    public bool HasTotalRow
    {
        get => tableProperties.LastRow?.Value ?? false;
        set => tableProperties.LastRow = value;
    }

    /// <inheritdoc/>
    public bool HasBandedRows
    {
        get => tableProperties.BandRow?.Value ?? false;
        set => tableProperties.BandRow = value;
    }

    /// <inheritdoc/>
    public bool HasFirstColumn
    {
        get => tableProperties.FirstColumn?.Value ?? false;
        set => tableProperties.FirstColumn = value;
    }

    /// <inheritdoc/>
    public bool HasLastColumn
    {
        get => tableProperties.LastColumn?.Value ?? false;
        set => tableProperties.LastColumn = value;
    }

    /// <inheritdoc/>
    public bool HasBandedColumns
    {
        get => tableProperties.BandColumn?.Value ?? false;
        set => tableProperties.BandColumn = value;
    }
}