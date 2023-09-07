using System;
using System.Globalization;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartPoint : IChartPoint
{
    private readonly C.NumericValue cNumericValue;

    internal ChartPoint(C.NumericValue cNumericValue)
    {
        this.cNumericValue = cNumericValue;
    }

    public double Value
    {
        get => this.ParseValue();
        set => this.UpdateValue(value);
    }

    private double ParseValue()
    {
        var cachedValue = double.Parse(this.cNumericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);

        return Math.Round(cachedValue, 2);
    }

    private void UpdateValue(double value)
    {
        this.cNumericValue!.Text = value.ToString(CultureInfo.InvariantCulture);
    }
}