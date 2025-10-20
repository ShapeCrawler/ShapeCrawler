using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartTitleAlignment : IChartTitleAlignment
{
    private readonly ChartPart chartPart;

    internal ChartTitleAlignment(ChartPart chartPart)
    {
        this.chartPart = chartPart;
    }

    public decimal CustomAngle
    {
        get => this.GetCustomAngle();
        set => this.SetCustomAngle(value);
    }

    private decimal GetCustomAngle()
    {
        var cChart = this.chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;
        if (cTitle == null)
        {
            return 0;
        }

        var cRichText = cTitle.GetFirstChild<C.ChartText>()?.GetFirstChild<C.RichText>();
        if (cRichText == null)
        {
            return 0;
        }

        var aBodyProperties = cRichText.GetFirstChild<A.BodyProperties>();
        if (aBodyProperties?.Rotation?.Value != null)
        {
            // OpenXML rotation angles are stored in units of 1/60,000th of a degree
            return aBodyProperties.Rotation.Value / 60000m;
        }

        return 0;
    }

    private void SetCustomAngle(decimal angle)
    {
        var cChart = this.chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;

        // Ensure title structure exists
        if (cTitle == null)
        {
            cTitle = new C.Title();
            cChart!.InsertAt(cTitle, 0);
        }

        var cChartText = cTitle.GetFirstChild<C.ChartText>();
        if (cChartText == null)
        {
            cChartText = new C.ChartText();
            cTitle.AppendChild(cChartText);
        }

        var cRichText = cChartText.GetFirstChild<C.RichText>();
        if (cRichText == null)
        {
            cRichText = new C.RichText();
            cChartText.AppendChild(cRichText);
            cRichText.Append(new A.BodyProperties());
            cRichText.Append(new A.ListStyle());
            
            // Add at least one paragraph with empty text to satisfy OpenXML schema
            var aParagraph = new A.Paragraph();
            aParagraph.Append(new A.Run(new A.Text(" ")));
            cRichText.Append(aParagraph);
        }

        var aBodyProperties = cRichText.GetFirstChild<A.BodyProperties>();
        if (aBodyProperties == null)
        {
            aBodyProperties = new A.BodyProperties();
            cRichText.InsertAt(aBodyProperties, 0);
        }

        // OpenXML rotation angles are stored in units of 1/60,000th of a degree
        var rotationInSixtyThousandths = (int)(angle * 60000m);
        aBodyProperties.Rotation = rotationInSixtyThousandths;
    }
}