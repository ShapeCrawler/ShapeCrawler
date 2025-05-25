using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Slides;

internal sealed class ChartCollection(SlidePart slidePart)
{
    public void AddPieChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        if (seriesName == null)
        {
            throw new ArgumentNullException(nameof(seriesName));
        }

        new SCSlidePart(slidePart).AddPieChart(x, y, width, height, categoryValues, seriesName);
    }

    internal void AddBarChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<string, double> categoryValues,
        string seriesName)
    {
        new SCSlidePart(slidePart).AddBarChart(x, y, width, height, categoryValues, seriesName);
    }

    internal void AddScatterChart(
        int x,
        int y,
        int width,
        int height,
        Dictionary<double, double> pointValues,
        string seriesName)
    {
        new SCSlidePart(slidePart).AddScatterChart(x, y, width, height, pointValues, seriesName);
    }

    internal void AddStackedColumnChart(
        int x,
        int y,
        int width,
        int height,
        IDictionary<string, IList<double>> categoryValues,
        IList<string> seriesNames)
    {
        new SCSlidePart(slidePart).AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);
    }
}