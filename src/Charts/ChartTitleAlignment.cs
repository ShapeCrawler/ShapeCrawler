using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace ShapeCrawler.Charts;

internal sealed class ChartTitleAlignment(ChartPart chartPart) : IChartTitleAlignment
{
    public decimal CustomAngle
    {
        get => this.GetCustomAngle();
        set => this.SetCustomAngle(value);
    }

    public decimal? X
    {
        get => this.GetX();
        set => this.SetX(value);
    }

    public decimal? Y
    {
        get => this.GetY();
        set => this.SetY(value);
    }

    private decimal GetCustomAngle()
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
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
            // Open XML rotation angles are stored in units of 1/60,000th of a degree
            return aBodyProperties.Rotation.Value / 60000m;
        }

        return 0;
    }

    private void SetCustomAngle(decimal angle)
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
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

    private decimal? GetX()
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;
        if (cTitle == null)
        {
            return null;
        }

        var cLayout = cTitle.GetFirstChild<C.Layout>();
        var cManualLayout = cLayout?.GetFirstChild<C.ManualLayout>();
        if (cManualLayout == null)
        {
            return null;
        }

        var cLeft = cManualLayout.GetFirstChild<C.Left>();
        return cLeft?.Val?.Value != null ? (decimal)cLeft.Val.Value : null;
    }

    private void SetX(decimal? value)
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;

        // Ensure title structure exists
        if (cTitle == null)
        {
            cTitle = new C.Title();
            cChart!.InsertAt(cTitle, 0);
        }

        var cLayout = cTitle.GetFirstChild<C.Layout>();
        if (cLayout == null)
        {
            cLayout = new C.Layout();
            cTitle.AppendChild(cLayout);
        }

        var cManualLayout = cLayout.GetFirstChild<C.ManualLayout>();
        
        if (value == null)
        {
            // Remove manual layout if setting to null (return to automatic positioning)
            cManualLayout?.Remove();
            if (cLayout.ChildElements.Count == 0)
            {
                cLayout.Remove();
            }

            return;
        }

        if (cManualLayout == null)
        {
            cManualLayout = new C.ManualLayout();
            cLayout.AppendChild(cManualLayout);
        }

        // Ensure LeftMode is set to factor
        var cLeftMode = cManualLayout.GetFirstChild<C.LeftMode>();
        if (cLeftMode == null)
        {
            cLeftMode = new C.LeftMode { Val = C.LayoutModeValues.Factor };
            cManualLayout.AppendChild(cLeftMode);
        }
        else
        {
            cLeftMode.Val = C.LayoutModeValues.Factor;
        }

        // Set the Left value
        var cLeft = cManualLayout.GetFirstChild<C.Left>();
        if (cLeft == null)
        {
            cLeft = new C.Left();
            cManualLayout.AppendChild(cLeft);
        }

        cLeft.Val = (double)value.Value;
    }

    private decimal? GetY()
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;
        if (cTitle == null)
        {
            return null;
        }

        var cLayout = cTitle.GetFirstChild<C.Layout>();
        var cManualLayout = cLayout?.GetFirstChild<C.ManualLayout>();
        if (cManualLayout == null)
        {
            return null;
        }

        var cTop = cManualLayout.GetFirstChild<C.Top>();
        return cTop?.Val?.Value != null ? (decimal)cTop.Val.Value : null;
    }

    private void SetY(decimal? value)
    {
        var cChart = chartPart.ChartSpace.GetFirstChild<C.Chart>();
        var cTitle = cChart?.Title;

        // Ensure title structure exists
        if (cTitle == null)
        {
            cTitle = new C.Title();
            cChart!.InsertAt(cTitle, 0);
        }

        var cLayout = cTitle.GetFirstChild<C.Layout>();
        if (cLayout == null)
        {
            cLayout = new C.Layout();
            cTitle.AppendChild(cLayout);
        }

        var cManualLayout = cLayout.GetFirstChild<C.ManualLayout>();
        
        if (value == null)
        {
            // Remove manual layout if setting to null (return to automatic positioning)
            cManualLayout?.Remove();
            if (cLayout.ChildElements.Count == 0)
            {
                cLayout.Remove();
            }

            return;
        }

        if (cManualLayout == null)
        {
            cManualLayout = new C.ManualLayout();
            cLayout.AppendChild(cManualLayout);
        }

        // Ensure TopMode is set to factor
        var cTopMode = cManualLayout.GetFirstChild<C.TopMode>();
        if (cTopMode == null)
        {
            cTopMode = new C.TopMode { Val = C.LayoutModeValues.Factor };
            cManualLayout.AppendChild(cTopMode);
        }
        else
        {
            cTopMode.Val = C.LayoutModeValues.Factor;
        }

        // Set the Top value
        var cTop = cManualLayout.GetFirstChild<C.Top>();
        if (cTop == null)
        {
            cTop = new C.Top();
            cManualLayout.AppendChild(cTop);
        }

        cTop.Val = (double)value.Value;
    }
}