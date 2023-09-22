namespace ShapeCrawler;

using C = DocumentFormat.OpenXml.Drawing.Charts;

/// <summary>
///     Represents a chart axes manager.
/// </summary>
public interface IAxesManager
{
    /// <summary>
    ///     Gets value axis. Returns <see langword="null"/> if chart has no value axis, e.g. pie chart.
    /// </summary>
    IAxis? ValueAxis { get; }
}

internal sealed class AxesManager : IAxesManager
{
    private readonly C.PlotArea cPlotArea;

    internal AxesManager(C.PlotArea cPlotArea)
    {
        this.cPlotArea = cPlotArea;
    }

    public IAxis? ValueAxis => this.GetValueAxis();

    private IAxis? GetValueAxis()
    {
        var cValueAxis = this.cPlotArea.GetFirstChild<C.ValueAxis>();
        return cValueAxis == null ? null : new Axis(cValueAxis.Scaling!);
    }
}
