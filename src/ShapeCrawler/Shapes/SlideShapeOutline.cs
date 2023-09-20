using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class SlideShapeOutline : IShapeOutline
{
    private readonly TypedOpenXmlCompositeElement sdkTypedOpenXmlCompositeElement;
    private readonly SlidePart sdkSlidePart;

    internal SlideShapeOutline(SlidePart sdkSlidePart, TypedOpenXmlCompositeElement sdkTypedOpenXmlCompositeElement)
    {
        this.sdkTypedOpenXmlCompositeElement = sdkTypedOpenXmlCompositeElement;
        this.sdkSlidePart = sdkSlidePart;
    }

    public double Weight
    {
        get => this.ParseWeight();
        set => this.UpdateWeight(value);
    }

    public string? HexColor
    {
        get => this.ParseHexColor();
        set => this.UpdateHexColor(value);
    }

    private void UpdateWeight(double points)
    {
        var aOutline = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.sdkTypedOpenXmlCompositeElement.AddAOutline();
        }

        aOutline.Width = new Int32Value(UnitConverter.PointToEmu(points));
    }
    
    private void UpdateHexColor(string? hex)
    {
        var aOutline = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.sdkTypedOpenXmlCompositeElement.AddAOutline();
        }

        var aSolidFill = aOutline.GetFirstChild<A.SolidFill>();
        aNoFill?.Remove();
        aSolidFill?.Remove();

        var aSrgbColor = new A.RgbColorModelHex { Val = hex };
        aSolidFill = new A.SolidFill(aSrgbColor);
        aOutline.Append(aSolidFill);
    }

    private double ParseWeight()
    {
        var width = this.sdkTypedOpenXmlCompositeElement.GetFirstChild<A.Outline>()?.Width;
        if (width is null)
        {
            return 0;
        }

        var widthEmu = width.Value;

        return UnitConverter.EmuToPoint(widthEmu);
    }

    private string? ParseHexColor()
    {
        var aSolidFill = this.sdkTypedOpenXmlCompositeElement
            .GetFirstChild<A.Outline>()?
            .GetFirstChild<A.SolidFill>();
        if (aSolidFill is null)
        {
            var defaultBlackHex = "000000";
            return defaultBlackHex;
        }

        var typeAndHex = HexParser.FromSolidFill(aSolidFill, this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster);
        
        return typeAndHex.Item2;
    }
}