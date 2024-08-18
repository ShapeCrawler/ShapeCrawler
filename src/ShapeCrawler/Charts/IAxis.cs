// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

using C = DocumentFormat.OpenXml.Drawing.Charts;

/// <summary>
///     Represents a chart axis.
/// </summary>
public interface IAxis
{
    /// <summary>
    ///     Gets or sets axis minimum value.
    /// </summary>
    double Minimum { get; set; }
    
    /// <summary>
    ///     Gets or sets axis maximum value.
    /// </summary>
    double Maximum { get; set; }
}

internal sealed record Axis : IAxis
{
    private const double DefaultMax = 6;
    private readonly C.Scaling cScaling;

    internal Axis(C.Scaling cScaling)
    {
        this.cScaling = cScaling;
    }

    public double Minimum
    {
        get => this.GetMinimum();
        set => this.cScaling.MinAxisValue = new C.MinAxisValue { Val = value };
    }

    public double Maximum
    {
        get => this.GetMaximum();
        set => this.cScaling.MaxAxisValue = new C.MaxAxisValue { Val = value };
    }

    private double GetMinimum()
    {
        var cMin = this.cScaling.MinAxisValue;
        
        return cMin == null ? 0 : cMin.Val!;
    }
    
    private double GetMaximum()
    {
        var cMax = this.cScaling.MaxAxisValue;
        
        return cMax == null ? DefaultMax : cMax.Val!;
    }
}