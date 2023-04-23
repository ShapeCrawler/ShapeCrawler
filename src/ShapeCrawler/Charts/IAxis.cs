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

internal class SCAxis : IAxis
{
    private const double DefaultMax = 6;
    private readonly C.Scaling cScaling;

    public SCAxis(C.Scaling cScaling)
    {
        this.cScaling = cScaling;
    }

    public double Minimum
    {
        get => this.GetMinimum();
        set => this.SetMinimum(value);
    }

    public double Maximum
    {
        get => this.GetMaximum();
        set => this.SetMaximum(value);
    }

    private void SetMaximum(double value)
    {
        this.cScaling.MaxAxisValue = new C.MaxAxisValue { Val = value };
    }

    private void SetMinimum(double value)
    {
        this.cScaling.MinAxisValue = new C.MinAxisValue { Val = value };
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