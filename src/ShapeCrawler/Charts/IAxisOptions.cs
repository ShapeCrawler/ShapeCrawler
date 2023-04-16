using ShapeCrawler.Charts;

namespace ShapeCrawler;

public interface IAxisOptions
{
    /// <summary>
    ///     Gets bounds. Returns <see langword="null"/> if bounds are not available, for example, for pie charts.
    /// </summary>
    IBounds? Bounds { get; }
}

internal class SCAxisOptions : IAxisOptions
{
    private readonly DocumentFormat.OpenXml.Drawing.Charts.Scaling cScaling;

    internal SCAxisOptions(DocumentFormat.OpenXml.Drawing.Charts.Scaling cScaling)
    {
        this.cScaling = cScaling;
    }

    public IBounds? Bounds => this.GetBounds();

    private IBounds GetBounds()
    {
        return new SCBounds(this.cScaling);
    }
}