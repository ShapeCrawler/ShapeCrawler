using DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler;

public interface IFormatAxis
{
    IAxisOptions AxisOptions { get; }
}

internal class SCFormatAxis : IFormatAxis
{
    private readonly Scaling cScaling;

    public SCFormatAxis(Scaling cScaling)
    {
        this.cScaling = cScaling;
    }

    public IAxisOptions AxisOptions => this.GetAxisOptions();

    private IAxisOptions GetAxisOptions()
    {
        return new SCAxisOptions(this.cScaling);
    }
}