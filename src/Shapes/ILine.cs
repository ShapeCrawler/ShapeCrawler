using ShapeCrawler.Shapes;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;

/// <summary>
///     Represents a line shape.
/// </summary>
public interface ILine
{
    /// <summary>
    ///    Gets the start point of the line.
    /// </summary>
    Point StartPoint { get; }

    /// <summary>
    ///     Gets the end point of the line.
    /// </summary>
    Point EndPoint { get; }
}

internal sealed class Line(P.ConnectionShape pConnectionShape, LineShape parentLineShape) : ILine
{
    public Geometry GeometryType
    {
        get => Geometry.Line;
        set => throw new SCException("It is not possible to set the geometry type for the chart shape.");
    }

    public Point StartPoint
    {
        get
        {
            var aTransform2D = pConnectionShape.GetFirstChild<P.ShapeProperties>()!.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            if (flipH && (parentLineShape.Height == 0 || flipV))
            {
                return new Point(parentLineShape.X, parentLineShape.Y);
            }

            if (flipH)
            {
                return new Point(parentLineShape.X + parentLineShape.Width, parentLineShape.Y);
            }

            return new Point(parentLineShape.X, parentLineShape.Y);
        }
    }

    public Point EndPoint
    {
        get
        {
            var aTransform2D = pConnectionShape.GetFirstChild<P.ShapeProperties>()!.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            if (parentLineShape.Width == 0)
            {
                return new Point(parentLineShape.X, parentLineShape.Height);
            }

            if (flipH && parentLineShape.Height == 0)
            {
                return new Point(parentLineShape.X - parentLineShape.Width, parentLineShape.Y);
            }

            if (flipV)
            {
                return new Point(parentLineShape.Width, parentLineShape.Height);
            }

            if (flipH)
            {
                return new Point(parentLineShape.X, parentLineShape.Height);
            }

            return new Point(parentLineShape.Width, parentLineShape.Y);
        }
    }
}