using System;
using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartPoint : IChartPoint
{
    private readonly ChartPart chartPart;
    private readonly C.NumericValue cNumericValue;
    private readonly string worksheetName;
    private readonly string address;

    internal ChartPoint(ChartPart chartPart, C.NumericValue cNumericValue, string worksheetName, string address)
    {
        this.chartPart = chartPart;
        this.cNumericValue = cNumericValue;
        this.worksheetName = worksheetName;
        this.address = address;
    }

    public double Value
    {
        get
        {
            var cachedValue = double.Parse(this.cNumericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);

            return Math.Round(cachedValue, 2);
        }

        set
        {
            this.cNumericValue.Text = value.ToString(CultureInfo.InvariantCulture);

            if (this.chartPart.EmbeddedPackagePart == null)
            {
                return;
            }

            new Workbook(this.chartPart.EmbeddedPackagePart).Sheet(this.worksheetName)
                .UpdateCell(this.address, value.ToString(CultureInfo.InvariantCulture));
        }
    }
}