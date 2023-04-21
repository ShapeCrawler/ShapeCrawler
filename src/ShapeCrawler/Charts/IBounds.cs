using System;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents axis bounds.
/// </summary>
public interface IBounds
{
    /// <summary>
    ///     Gets or sets the minimum value of the axis.
    /// </summary>
    double Minimum { get; set; }

    /// <summary>
    ///     Gets or sets the maximum value of the axis.
    /// </summary>
    double Maximum { get; set; }
}

internal class SCBounds : IBounds
{
    private const double DefaultMax = 6;
    private readonly DocumentFormat.OpenXml.Drawing.Charts.Scaling cScaling;

    public SCBounds(DocumentFormat.OpenXml.Drawing.Charts.Scaling cScaling)
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
        throw new NotImplementedException();
    }

    private double GetMaximum()
    {
        var cMax = this.cScaling.MaxAxisValue;
        return cMax == null ? DefaultMax : cMax.Val!;
    }

    private void SetMinimum(double value)
    {
        throw new NotImplementedException();
    }

    private double GetMinimum()
    {
        var cMin = this.cScaling.MinAxisValue;
        return cMin == null ? 0 : cMin.Val!;
    }
}