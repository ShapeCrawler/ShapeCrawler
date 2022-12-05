using DocumentFormat.OpenXml;
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
        if (aOutline is null)
        {
            aOutline = new A.Outline
            {
                Width = new Int32Value()
            };
            pShapeProperties.AppendChild(aOutline);
        }
        
        aOutline.Width!.Value = UnitConverter.PointToEmu(points);
    }

    private double GetWeight()
    {
        var aOutline = this.parentAutoShape.PShapeTreesChild.GetFirstChild<P.ShapeProperties>() !.GetFirstChild<A.Outline>();
        if (aOutline is null)
        {
            return 0;
        }

        var widthEmu = aOutline.Width!.Value;

        return UnitConverter.EmuToPoint(widthEmu);
    }
}