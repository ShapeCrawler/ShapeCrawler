using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartTitle : IChartTitle
{
    private readonly ChartPart chartPart;
    private readonly Func<ChartType> getChartType;
    private readonly Func<ISeriesCollection> getSeriesCollection;

    internal ChartTitle(ChartPart chartPart, Func<ChartType> getChartType, Func<ISeriesCollection> getSeriesCollection)
    {
        this.chartPart = chartPart;
        this.getChartType = getChartType;
        this.getSeriesCollection = getSeriesCollection;
    }

    public string? Text
    {
        get => this.GetTitleText();
        set => this.UpdateTitleText(value);
    }

    public string FontColor
    {
        get => this.GetFontColor();
        set => this.SetFontColor(value);
    }

    public static implicit operator string?(ChartTitle? title) => title?.Text;

    public override string? ToString() => this.Text;

    private static bool TryGetStaticTitle(C.ChartText? chartText, ChartType chartType, out string? staticTitle)
    {
        staticTitle = null;
        if (chartText == null)
        {
            return false;
        }

        if (chartType == ChartType.Combination && chartText.RichText != null)
        {
            var texts = chartText.RichText.Descendants<A.Text>().Select(t => t.Text);
            staticTitle = string.Concat(texts);
            return true;
        }

        var rRich = chartText.RichText;
        if (rRich != null)
        {
            var texts = rRich.Descendants<A.Text>().Select(t => t.Text);
            staticTitle = string.Concat(texts);
            return true;
        }

        return false;
    }

    private string GetFontColor()
    {
        var cChart = this.chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;
        if (cTitle == null)
        {
            return "000000"; // default black
        }

        var aText = cTitle.Descendants<A.Text>().FirstOrDefault();
        if (aText == null)
        {
            return "000000";
        }

        var aRunProperties = aText.Parent?.GetFirstChild<A.RunProperties>();
        var aSolidFill = aRunProperties?.GetFirstChild<A.SolidFill>();
        var rgbColor = aSolidFill?.RgbColorModelHex;

        if (rgbColor?.Val?.Value != null)
        {
            return rgbColor.Val.Value;
        }

        return "000000";
    }

    private void SetFontColor(string hex)
    {
        var cChart = this.chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cRichText = cChart?.Title?.GetFirstChild<C.ChartText>()?.GetFirstChild<C.RichText>();
        if (cRichText is null)
        {
            return;
        }

        // Process hex value
        hex = hex.StartsWith("#", StringComparison.Ordinal) ? hex[1..] : hex;
        if (hex.Length == 8)
        {
            hex = hex[..6];
        }

        // Apply color to all existing runs
        foreach (var aParagraph in cRichText.Elements<A.Paragraph>())
        {
            foreach (var aRun in aParagraph.Elements<A.Run>())
            {
                var aRunProperties = aRun.GetFirstChild<A.RunProperties>();
                if (aRunProperties == null)
                {
                    aRunProperties = new A.RunProperties();
                    aRun.InsertAt(aRunProperties, 0);
                }

                // Remove existing solid fill
                var existingSolidFill = aRunProperties.GetFirstChild<A.SolidFill>();
                existingSolidFill?.Remove();

                // Add new solid fill with color
                var aSolidFill = new A.SolidFill();
                var rgbColorModelHex = new A.RgbColorModelHex { Val = hex };
                aSolidFill.AppendChild(rgbColorModelHex);
                aRunProperties.InsertAt(aSolidFill, 0);
            }
        }
    }

    private string? GetTitleText()
    {
        var cTitle = this.chartPart.ChartSpace.GetFirstChild<C.Chart>()!.Title;
        if (cTitle == null)
        {
            return null;
        }

        var cChartText = cTitle.ChartText;
        var chartType = this.getChartType();

        // Try static title
        if (TryGetStaticTitle(cChartText!, chartType, out var staticTitle))
        {
            return staticTitle;
        }

        // Dynamic title
        if (cChartText != null)
        {
            var stringPoint = cChartText.Descendants<C.StringPoint>().FirstOrDefault();
            if (stringPoint != null)
            {
                return stringPoint.InnerText;
            }
        }

        // PieChart uses only one series for view
        if (chartType == ChartType.PieChart)
        {
            return this.getSeriesCollection().FirstOrDefault()?.Name;
        }

        return null;
    }

    private void UpdateTitleText(string? value)
    {
        // Delegate to the existing SetTitle method in Chart
        var cChart = this.chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var cTitle = cChart.Title;

        if (string.IsNullOrEmpty(value))
        {
            cTitle?.Remove();
            var plotArea = cChart.PlotArea!;
            var autoTitleDeleted = plotArea.GetFirstChild<C.AutoTitleDeleted>();
            if (autoTitleDeleted == null)
            {
                plotArea.InsertAt(new C.AutoTitleDeleted { Val = true }, 0);
            }
            else
            {
                autoTitleDeleted.Val = true;
            }

            return;
        }

        var autoTitleDeletedCheck = cChart.PlotArea!.GetFirstChild<C.AutoTitleDeleted>();
        if (autoTitleDeletedCheck != null)
        {
            autoTitleDeletedCheck.Val = false;
        }

        if (cTitle == null)
        {
            cTitle = new C.Title();
            cChart.InsertAt(cTitle, 0);
        }

        var cChartText = cTitle.GetFirstChild<C.ChartText>() ?? cTitle.AppendChild(new C.ChartText());

        var cRichText = cChartText.GetFirstChild<C.RichText>();
        if (cRichText == null)
        {
            cRichText = cChartText.AppendChild(new C.RichText());
            cRichText.Append(new A.BodyProperties());
            cRichText.Append(new A.ListStyle());
        }

        cRichText.RemoveAllChildren<A.Paragraph>();
        var aParagraph = cRichText.AppendChild(new A.Paragraph());
        aParagraph.Append(new A.Run(new A.Text(value!)));

        if (cTitle.Layout == null)
        {
            cTitle.Append(new C.Layout());
        }

        var cOverlay = cTitle.GetFirstChild<C.Overlay>() ?? cTitle.AppendChild(new C.Overlay());

        cOverlay.Val = false;
    }
}