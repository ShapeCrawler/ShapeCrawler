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
            var cXValues = this.FirstSeries().GetFirstChild<C.XValues>()!;

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

                return [.. cachedPointValues];
            }

            return [.. new Workbook(chartPart.EmbeddedPackagePart!).FormulaValues(cXValues.NumberReference.Formula!.Text)];
        }
    }

    public double Minimum
    {
        get
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!.Scaling!;
            var cMin = cScaling.MinAxisValue;
            
            return cMin == null ? 0 : cMin.Val!;
        }

        set
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!.Scaling!;
            cScaling.MinAxisValue = new C.MinAxisValue { Val = value };
        }
    }

    public double Maximum
    {
        get
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!.Scaling!;
            var cMax = cScaling.MaxAxisValue;
            const double defaultMax = 6;
            
            return cMax == null ? defaultMax : cMax.Val!;
        }

        set
        {
            var cScaling = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!.GetFirstChild<C.ValueAxis>()!.Scaling!;
            cScaling.MaxAxisValue = new C.MaxAxisValue { Val = value };
        }
    }

    private OpenXmlElement FirstSeries()
    {
        var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>() !.PlotArea!;
        var cXCharts = plotArea.Where(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        
        return cXCharts.First().ChildElements
            .First(e => e.LocalName.Equals("ser", StringComparison.Ordinal));
    }
}