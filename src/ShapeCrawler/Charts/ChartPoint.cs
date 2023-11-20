using System;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartPoint : IChartPoint
{
    private readonly ChartPart sdkChartPart;
    private readonly C.NumericValue cNumericValue;
    private readonly string sheet;
    private readonly string address;

    internal ChartPoint(ChartPart sdkChartPart, C.NumericValue cNumericValue, string sheet, string address)
    {
        this.sdkChartPart = sdkChartPart;
        this.cNumericValue = cNumericValue;
        this.sheet = sheet;
        this.address = address;
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
        this.cNumericValue.Text = value.ToString(CultureInfo.InvariantCulture);

        if (this.sdkChartPart.EmbeddedPackagePart == null)
        {
            return;
        }

        new ExcelBook(this.sdkChartPart).Sheet(this.sheet)
            .UpdateCell(this.address, value.ToString(CultureInfo.InvariantCulture));
    }
}