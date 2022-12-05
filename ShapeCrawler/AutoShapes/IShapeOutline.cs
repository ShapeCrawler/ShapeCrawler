using ShapeCrawler.Statics;

// ReSharper disable once CheckNamespace
namespace ShapeCrawler;

/// <summary>
///     Represents a shape outline.
/// </summary>
public interface IShapeOutline
{
    /// <summary>
    ///     Gets outline weight in points.
    /// </summary>
    double Weight { get; }
}

internal class SCShapeOutline : IShapeOutline
{
    private readonly SlideAutoShape parentAutoShape;

    internal SCShapeOutline(SlideAutoShape parentAutoShape)
    {
        this.parentAutoShape = parentAutoShape;
    }

    public double Weight => this.GetWeight();

    private double GetWeight()
    {
        var aOutline = this.parentAutoShape.PShapeTreesChild.GetFirstChild<DocumentFormat.OpenXml.Presentation.ShapeProperties>() !.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();
        if (aOutline is null)
        {
            return 0;
        }

        var widthEmu = aOutline.Width!.Value;

        return UnitConverter.EmuToPoint(widthEmu);
    }
}