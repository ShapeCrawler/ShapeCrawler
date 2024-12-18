namespace ShapeCrawler.Tables;

/// <summary>
///     Represents table style options.
/// </summary>
public interface ITableStyleOptions
{
    /// <summary>
    ///     Gets a value indicating whether table has header row.
    /// </summary>
    public bool HasHeaderRow { get; }
    
    /// <summary>
    ///     Gets a value indicating whether table has total row.
    /// </summary>
    public bool HasTotalRow { get; }
    
    /// <summary>
    ///     Gets a value indicating whether table has banded rows.
    /// </summary>
    public bool HasBandedRows { get; }
    
    /// <summary>
    ///     Gets a value indicating whether table has first column.
    /// </summary>
    public bool HasFirstColumn { get; }
    
    /// <summary>
    ///     Gets a value indicating whether table has last column.
    /// </summary>
    public bool HasLastColumn { get; }
    
    /// <summary>
    ///     Gets a value indicating whether table has banded columns.
    /// </summary>
    public bool HasBandedColumns { get; }
}

/// <summary>
///    Represents table style options.
/// </summary>
/// <param name="hasHeaderRow"></param>
/// <param name="hasTotalRow"></param>
/// <param name="hasBandedRows"></param>
/// <param name="hasFirstColumn"></param>
/// <param name="hasLastColumn"></param>
/// <param name="hasBandedColumns"></param>
public class TableStyleOptions(
    bool hasHeaderRow = false,
    bool hasTotalRow = false,
    bool hasBandedRows = false,
    bool hasFirstColumn = false,
    bool hasLastColumn = false,
    bool hasBandedColumns = false)
    : ITableStyleOptions
{
    /// <inheritdoc/>
    public bool HasHeaderRow { get; } = hasHeaderRow;
    
    /// <inheritdoc/>
    public bool HasTotalRow { get; } = hasTotalRow;
    
    /// <inheritdoc/>
    public bool HasBandedRows { get; } = hasBandedRows;
    
    /// <inheritdoc/>
    public bool HasFirstColumn { get; } = hasFirstColumn;
    
    /// <inheritdoc/>
    public bool HasLastColumn { get; } = hasLastColumn;
    
    /// <inheritdoc/>
    public bool HasBandedColumns { get; } = hasBandedColumns;
}