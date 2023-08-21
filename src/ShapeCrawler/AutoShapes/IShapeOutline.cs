using DocumentFormat.OpenXml;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.Drawing;
using ShapeCrawler.Extensions;
using ShapeCrawler.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape outline.
/// </summary>
public interface IShapeOutline
{
    /// <summary>
    ///     Gets or sets outline weight in points.
    /// </summary>
    double Weight { get; set; }

    /// <summary>
    ///     Gets or sets color in hexadecimal format. Returns <see langword="null"/> if outline is not filled.
    /// </summary>
    string? HexColor { get; set; }
}

internal sealed class ShapeOutline : IShapeOutline
{
    private readonly SlideMaster slideMaster;
    private readonly P.ShapeProperties pShapeProperties;

    internal ShapeOutline(SlideMaster slideMaster, P.ShapeProperties pShapeProperties)
    {
        this.slideMaster = slideMaster;
        this.pShapeProperties = pShapeProperties;
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
        var aOutline = pShapeProperties.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = pShapeProperties.AddAOutline();
        }

        aOutline.Width = new Int32Value(UnitConverter.PointToEmu(points));
    }
    
    private void UpdateHexColor(string? hex)
    {
        var aOutline = this.pShapeProperties.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = pShapeProperties.AddAOutline();
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

        var typeAndHex = HexParser.FromSolidFill(aSolidFill, this.slideMaster);
        
        return typeAndHex.Item2;
    }
}