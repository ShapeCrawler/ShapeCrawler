#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Bar chart type.
/// </summary>
public enum BarChartType
{
    /// <summary>
    ///     Clustered bar chart.
    /// </summary>
    ClusteredBar,
    
    /// <summary>
    ///     Stacked bar chart.
    /// </summary>
    Stacked,
    
    /// <summary>
    ///     100% stacked bar chart.
    /// </summary>
    Stacked100Percent,
    
    /// <summary>
    ///     Clustered 3D bar chart.
    /// </summary>
    Clustered3D,
    
    /// <summary>
    ///     Stacked 3D bar chart.
    /// </summary>
    Stacked3D,
    
    /// <summary>
    ///     100% stacked 3D bar chart.
    /// </summary>
    Stacked100Percent3D
}