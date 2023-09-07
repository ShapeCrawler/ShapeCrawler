using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal sealed class SlideShapeOutline : IShapeOutline
{
    private readonly P.ShapeProperties pShapeProperties;
    private readonly SlidePart sdkSlidePart;

    internal SlideShapeOutline(SlidePart sdkSlidePart, P.ShapeProperties pShapeProperties)
    {
        this.pShapeProperties = pShapeProperties;
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
        var aOutline = this.pShapeProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();
        var aNoFill = aOutline?.GetFirstChild<DocumentFormat.OpenXml.Drawing.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.pShapeProperties.AddAOutline();
        }

        aOutline.Width = new Int32Value(UnitConverter.PointToEmu(points));
    }
    
    private void UpdateHexColor(string? hex)
    {
        var aOutline = this.pShapeProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();
        var aNoFill = aOutline?.GetFirstChild<DocumentFormat.OpenXml.Drawing.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = this.pShapeProperties.AddAOutline();
        }

        var aSolidFill = aOutline.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();
        aNoFill?.Remove();
        aSolidFill?.Remove();

        var aSrgbColor = new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = hex };
        aSolidFill = new DocumentFormat.OpenXml.Drawing.SolidFill(aSrgbColor);
        aOutline.Append(aSolidFill);
    }

    private double ParseWeight()
    {
        var width = this.pShapeProperties.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()?.Width;
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
            .GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>()?
            .GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();
        if (aSolidFill is null)
        {
            return null;
        }

        var typeAndHex = HexParser.FromSolidFill(aSolidFill, this.sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster!);
        
        return typeAndHex.Item2;
    }
}