using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using ShapeCrawler.Statics;
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
}

internal class SCShapeOutline : IShapeOutline
{
    private readonly SlideAutoShape parentAutoShape;

    internal SCShapeOutline(SlideAutoShape parentAutoShape)
    {
        this.parentAutoShape = parentAutoShape;
    }

    public double Weight
    {
        get => this.GetWeight();
        set => this.SetWeight(value);
    }

    private void SetWeight(double points)
    {
        var pShapeProperties = this.parentAutoShape.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !;
        var aOutline = pShapeProperties.GetFirstChild<A.Outline>();
        var aNoFill = aOutline?.GetFirstChild<A.NoFill>();

        if (aOutline == null || aNoFill != null)
        {
            aOutline = pShapeProperties.AddAOutline();
        }

        aOutline.Width = new Int32Value(UnitConverter.PointToEmu(points));
    }

    private double GetWeight()
    {
        var width = this.parentAutoShape.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !.GetFirstChild<A.Outline>()?.Width;
        if (width is null)
        {
            return 0;
        }

        var widthEmu = width.Value;

        return UnitConverter.EmuToPoint(widthEmu);
    }
}