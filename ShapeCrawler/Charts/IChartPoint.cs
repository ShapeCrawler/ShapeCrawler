using System;
using System.Globalization;
using ShapeCrawler.Charts;
using ShapeCrawler.Exceptions;
using C = DocumentFormat.OpenXml.Drawing.Charts;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a chart point.
/// </summary>
public interface IChartPoint
{
    /// <summary>
    ///     Gets or sets chart point value.
    /// </summary>
    public double Value { get; set; }
}

internal class ChartPoint : IChartPoint
{
    private readonly string address;
    private readonly SCChart parentChart;
    private readonly string sheetName;
    private readonly C.NumericValue? cNumericValue;
    private readonly ChartWorkbook? workbook;

    internal ChartPoint(SCChart parentChart, string sheetName, string address, C.NumericValue? cNumericValue)
        : this(parentChart, sheetName, address)
    {
        this.cNumericValue = cNumericValue;
    }

    private ChartPoint(SCChart chart, string sheetName, string address)
    {
        this.parentChart = chart;
        this.sheetName = sheetName;
        this.address = address;
        this.workbook = chart.ChartWorkbook;
    }

    public double Value
    {
        get
        {
            var context = $"Chart type:\t{this.parentChart.Type.ToString()}";
            ErrorHandler.Execute(this.GetValue, context, out var result);

            return result;
        }

        set
        {
            var context = $"Chart type:\t{this.parentChart.Type.ToString()}";
            ErrorHandler.Execute(() => this.UpdateValue(value), context);
        }
    }

    private double GetValue()
    {
        // From cache
        if (this.cNumericValue != null)
        {
            var cachedValue = double.Parse(this.cNumericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
            return Math.Round(cachedValue, 2);
        }

        // From spreadsheet
        var xCell = this.workbook!.GetXCell(this.sheetName, this.address);
        var sheetValue = xCell.InnerText.Length == 0 ? 0 : double.Parse(xCell.InnerText, CultureInfo.InvariantCulture.NumberFormat);

        return sheetValue;
    }

    private void UpdateValue(double value)
    {
        if (this.cNumericValue != null)
        {
            this.cNumericValue.Text = value.ToString(CultureInfo.InvariantCulture);
        }

        if (this.workbook == null)
        {
            // Chart can have Linked file instead of Embedded. This Linked file can be removed
            return;
        }

        this.workbook.UpdateCell(this.sheetName, this.address, value);
    }
}