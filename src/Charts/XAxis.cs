using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal class XAxis(ChartPart chartPart) : IXAxis
{
    public double[] Values
    {
        get
        {
            var cXValues = FirstSeries().GetFirstChild<C.XValues>()!;

            if (cXValues.NumberReference!.NumberingCache != null)
            {
                var cNumericValues = cXValues.NumberReference.NumberingCache.Descendants<C.NumericValue>();
                var cachedPointValues = new List<double>(cNumericValues.Count());
                foreach (var numericValue in cNumericValues)
                {
                    var number = double.Parse(numericValue.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    var roundNumber = Math.Round(number, 1);
                    cachedPointValues.Add(roundNumber);
                }

                return cachedPointValues.ToArray();
            }

            return new Spreadsheet(chartPart).FormulaValues(cXValues.NumberReference.Formula!.Text).ToArray();
        }
    }

    public int Minimum { get; set; }
    public int Maximum { get; set; }

    private OpenXmlElement FirstSeries()
    {
        var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        
        return cXCharts.First().ChildElements
            .First(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
    }
}