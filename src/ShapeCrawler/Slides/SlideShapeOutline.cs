using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Units;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Slides;

internal sealed class SlideShapeOutline : IShapeOutline
{
    private readonly OpenXmlPart openXmlPart;
    private readonly OpenXmlCompositeElement openXmlCompositeElement;

    internal SlideShapeOutline(OpenXmlPart openXmlPart, OpenXmlCompositeElement openXmlCompositeElement)
    {
        this.openXmlPart = openXmlPart;
        this.openXmlCompositeElement = openXmlCompositeElement;
    }

    public decimal Weight
    {
        get => this.ParseWeight();
        set => this.UpdateWeight(value);
    }

    /// <inheritdoc/>
    public string? HexColor => this.ParseHexColor();

    /// <inheritdoc/>
    public void SetHexColor(string value) => this.UpdateFill(new A.SolidFill(new A.RgbColorModelHex { Val = value }));

    /// <inheritdoc/>
    public void SetNoOutline() => this.UpdateFill(new A.NoFill());

    private void UpdateWeight(decimal points)
    {
        var aOutline = this.openXmlCompositeElement.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.openXmlCompositeElement.AddAOutline();
        }

        aOutline.Width = new Int32Value((int)UnitConverter.PointToEmu(points));
    }

    private void UpdateFill(OpenXmlElement child)
    {
        // Ensure there is an outline
        var aOutline = this.openXmlCompositeElement.GetFirstChild<A.Outline>();
        if (aOutline is null)
        {
            aOutline = new A.Outline();
            this.openXmlCompositeElement.AppendChild(aOutline);
        }

        // Remove any explicit existing kinds of outline
        aOutline.RemoveAllChildren();

        // Set the new child value
        aOutline.AppendChild(child);
    }

    private decimal ParseWeight()
    {
        var width = this.openXmlCompositeElement.GetFirstChild<A.Outline>()?.Width;
        if (width is null)
        {
            return 0;
        }

        var widthEmu = width.Value;

        return UnitConverter.EmuToPoint(widthEmu);
    }

    private string? ParseHexColor()
    {
        var aSolidFill = this.openXmlCompositeElement
            .GetFirstChild<A.Outline>()?
            .GetFirstChild<A.SolidFill>();
        if (aSolidFill is null)
        {
            return null;
        }

        var pSlideMaster = this.openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster,
            SlideLayoutPart sdkSlideLayoutPart => sdkSlideLayoutPart.SlideMasterPart!.SlideMaster,
            _ => ((SlideMasterPart)this.openXmlPart).SlideMaster
        };
        var typeAndHex = HexParser.FromSolidFill(aSolidFill, pSlideMaster);

        return typeAndHex.Item2;
    }
}