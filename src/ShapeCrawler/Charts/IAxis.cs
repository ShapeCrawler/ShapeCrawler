using System;
using System.Linq;

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
    private readonly C.PlotArea cPlotArea;

    public SCAxis(C.PlotArea cPlotArea)
    {
        this.cPlotArea = cPlotArea;
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

    private void SetMinimum(double value)
    {
        throw new NotImplementedException();
    }

    private double GetMinimum()
    {
        var cScaling = this.cPlotArea.Descendants<C.Scaling>().First();
        var cMin = cScaling.MinAxisValue;
        
        return cMin == null ? 0 : cMin.Val!;
    }
    
    private double GetMaximum()
    {
        var cScaling = this.cPlotArea.Descendants<C.Scaling>().First();
        var cMax = cScaling.MaxAxisValue;
        
        return cMax == null ? DefaultMax : cMax.Val!;
    }
}