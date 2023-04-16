using System;

namespace ShapeCrawler.Charts;

public interface IBounds
{
    double Minimum { get; set; }
}

internal class SCBounds : IBounds
{
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