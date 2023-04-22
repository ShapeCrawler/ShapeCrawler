namespace ShapeCrawler;

using C = DocumentFormat.OpenXml.Drawing.Charts;

/// <summary>
///     Represents a chart axes manager.
/// </summary>
public interface IAxesManager
{
    /// <summary>
    ///     Gets horizintal axis.
    /// </summary>
    IAxis HorizontalAxis { get; }
}

internal class SCAxesManager : IAxesManager
{
    private readonly C.PlotArea cPlotArea;

    public SCAxesManager(C.PlotArea cPlotArea)
    {
        this.cPlotArea = cPlotArea;
    }

    public IAxis HorizontalAxis => new SCAxis(this.cPlotArea);
}
