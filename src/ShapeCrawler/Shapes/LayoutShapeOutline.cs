using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.AutoShapes;

internal sealed class LayoutShapeOutline : IShapeOutline
{
    private readonly P.ShapeProperties pShapeProperties;
    private readonly SlideLayoutPart sdkSlideLayoutPart;

    internal LayoutShapeOutline(P.ShapeProperties pShapeProperties, SlideLayoutPart sdkSlideLayoutPart)
    {
        this.pShapeProperties = pShapeProperties;
        this.sdkSlideLayoutPart = sdkSlideLayoutPart;
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
        var aOutline = this.pShapeProperties.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.pShapeProperties.AddAOutline();
        }

        aOutline.Width = new Int32Value(UnitConverter.PointToEmu(points));
    }
    
    private void UpdateHexColor(string? hex)
    {
        var aOutline = this.pShapeProperties.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.pShapeProperties.AddAOutline();
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
        var width = this.pShapeProperties.GetFirstChild<A.Outline>()?.Width;
        if (width is null)
        {
            return 0;
        }

        var widthEmu = width.Value;

        return UnitConverter.EmuToPoint(widthEmu);
    }

    private string? ParseHexColor()
    {
        var aSolidFill = this.pShapeProperties
            .GetFirstChild<A.Outline>()?
            .GetFirstChild<A.SolidFill>();
        if (aSolidFill is null)
        {
            return null;
        }

        var typeAndHex = HexParser.FromSolidFill(aSolidFill, this.sdkSlideLayoutPart.SlideMasterPart!.SlideMaster);
        
        return typeAndHex.Item2;
    }
}