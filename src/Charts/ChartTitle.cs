using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartTitle(ChartPart chartPart, ChartType chartType, ISeriesCollection seriesCollection, ChartTitleAlignment alignment) : IChartTitle
{
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

    public int FontSize
    {
        get => this.GetFontSize();
        set => this.SetFontSize(value);
    }

    public IChartTitleAlignment Alignment => alignment;

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
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
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
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
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
        foreach (var aRun in cRichText.Elements<A.Paragraph>().SelectMany(p => p.Elements<A.Run>()))
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

    private string? GetTitleText()
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var cTitle = cChart.Title;
        
        if (cTitle == null)
        {
            // Check if title was explicitly deleted
            var autoTitleDeleted = cChart.PlotArea?.GetFirstChild<C.AutoTitleDeleted>();
            if (autoTitleDeleted?.Val?.Value == true)
            {
                return null;
            }
            
            // PieChart uses only one series for view when no title is set
            if (chartType == ChartType.PieChart)
            {
                return seriesCollection.FirstOrDefault()?.Name;
            }
            
            return null;
        }

        var cChartText = cTitle.ChartText;

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
            return seriesCollection.FirstOrDefault()?.Name;
        }

        return null;
    }

    private void UpdateTitleText(string? value)
    {
        // Delegate to the existing SetTitle method in Chart
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
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

    private int GetFontSize()
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;
        if (cTitle == null)
        {
            return 18; // default font size
        }

        var aRunProperties = cTitle.Descendants<A.RunProperties>().FirstOrDefault();
        if (aRunProperties?.FontSize?.Value != null)
        {
            return aRunProperties.FontSize.Value / 100;
        }

        return 18; // default font size
    }

    private void SetFontSize(int fontSize)
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
        var cTitle = cChart.Title;

        // Ensure title structure exists
        if (cTitle == null)
        {
            cTitle = new C.Title();
            cChart.InsertAt(cTitle, 0);
        }

        var cChartText = cTitle.GetFirstChild<C.ChartText>() ?? cTitle.AppendChild(new C.ChartText());

        var cRichText = cChartText.GetFirstChild<C.RichText>();
        if (cRichText == null)
        {
            // Create title structure with default text if it doesn't exist
            var currentText = this.Text ?? "Chart Title";
            cRichText = cChartText.AppendChild(new C.RichText());
            cRichText.Append(new A.BodyProperties());
            cRichText.Append(new A.ListStyle());
            var aParagraph = cRichText.AppendChild(new A.Paragraph());
            aParagraph.Append(new A.Run(new A.Text(currentText)));
        }

        var fontSizeInHundredthsOfPoint = fontSize * 100;

        // Apply font size to all existing runs
        foreach (var aRun in cRichText.Elements<A.Paragraph>().SelectMany(p => p.Elements<A.Run>()))
        {
            var aRunProperties = aRun.GetFirstChild<A.RunProperties>();
            if (aRunProperties == null)
            {
                aRunProperties = new A.RunProperties();
                aRun.InsertAt(aRunProperties, 0);
            }

            aRunProperties.FontSize = fontSizeInHundredthsOfPoint;
        }

        if (cTitle.Layout == null)
        {
            cTitle.Append(new C.Layout());
        }

        var cOverlay = cTitle.GetFirstChild<C.Overlay>() ?? cTitle.AppendChild(new C.Overlay());
        cOverlay.Val = false;
    }
}