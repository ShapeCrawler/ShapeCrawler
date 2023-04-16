using System.Linq;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler;

/// <summary>
///     Represents format of chart axis.
/// </summary>
public interface IFormatAxis
{
    /// <summary>
    ///     Gets axis options.
    /// </summary>
    IAxisOptions AxisOptions { get; }
}

internal class SCFormatAxis : IFormatAxis
{
    private readonly C.PlotArea cPlotArea;

    public SCFormatAxis(C.PlotArea cPlotArea)
    {
        this.cPlotArea = cPlotArea;
    }

    public IAxisOptions AxisOptions => this.GetAxisOptions();

    private IAxisOptions GetAxisOptions()
    {
        var cScaling = this.cPlotArea.Descendants<C.Scaling>().First();
        return new SCAxisOptions(cScaling);
    }
}