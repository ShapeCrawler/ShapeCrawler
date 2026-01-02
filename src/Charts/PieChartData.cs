using System.Collections.Generic;

namespace ShapeCrawler.Charts;

internal sealed class PieChartData
{
    public List<double> Values { get; } = [];

    public List<string> Categories { get; } = [];

    public string Title { get; set; } = string.Empty;

    public bool ShowLegend { get; set; }

    public bool ShowDataLabels { get; set; }
}