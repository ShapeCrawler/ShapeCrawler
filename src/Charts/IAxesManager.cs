#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

using C = DocumentFormat.OpenXml.Drawing.Charts;


internal sealed record AxesManager
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
