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
    private readonly P.ConnectionShape connectionShape = pConnectionShape;
    private readonly LineShape lineShape = parentLineShape;

    public Geometry GeometryType
    {
        get => Geometry.Line;
    }

    public Point StartPoint
    {
        get
        {
            var aTransform2D = this.connectionShape.GetFirstChild<P.ShapeProperties>()!.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            var startX = flipH ? this.lineShape.X + this.lineShape.Width : this.lineShape.X;
            var startY = flipV ? this.lineShape.Y + this.lineShape.Height : this.lineShape.Y;
            return new Point(startX, startY);
        }
    }

    public Point EndPoint
    {
        get
        {
            var aTransform2D = this.connectionShape.GetFirstChild<P.ShapeProperties>()!.Transform2D!;
            var horizontalFlip = aTransform2D.HorizontalFlip?.Value;
            var flipH = horizontalFlip != null && horizontalFlip.Value;
            var verticalFlip = aTransform2D.VerticalFlip?.Value;
            var flipV = verticalFlip != null && verticalFlip.Value;

            var endX = flipH ? this.lineShape.X : this.lineShape.X + this.lineShape.Width;
            var endY = flipV ? this.lineShape.Y : this.lineShape.Y + this.lineShape.Height;
            return new Point(endX, endY);
        }
    }
}