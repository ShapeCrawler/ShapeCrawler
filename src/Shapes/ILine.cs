using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using ShapeCrawler.Shapes;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

#pragma warning disable IDE0130
namespace ShapeCrawler;
#pragma warning restore IDE0130

/// <summary>
///     Represents a line shape.
/// </summary>
public interface ILine
{
    /// <summary>
    ///     Gets the PowerPoint line type.
    /// </summary>
    LineType Type { get; }

    /// <summary>
    ///     Gets or sets the start point of the line.
    /// </summary>
    Point StartPoint { get; set; }

    /// <summary>
    ///     Gets or sets the end point of the line.
    /// </summary>
    Point EndPoint { get; set; }

    /// <summary>
    ///     Gets the points that define the line path. For connector lines this contains the start and end points.
    /// </summary>
    IReadOnlyList<Point> Points { get; }
}

internal sealed class Line(OpenXmlElement shapeElement, LineShape parentLineShape) : ILine
{
    private readonly OpenXmlElement shapeElement = shapeElement;
    private readonly LineShape lineShape = parentLineShape;

    public LineType Type
    {
        get
        {
            var preset = this.ShapeProperties.GetFirstChild<PresetGeometry>()?.Preset?.InnerText;
            if (preset is not null)
            {
                var outline = this.ShapeProperties.GetFirstChild<Outline>();
                return LineTypeMapping.FromOpenXml(
                    preset,
                    outline?.GetFirstChild<HeadEnd>()?.Type?.Value,
                    outline?.GetFirstChild<TailEnd>()?.Type?.Value);
            }

            var name = this.shapeElement switch
            {
                P.Shape pShape => pShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                P.ConnectionShape pConnectionShape => pConnectionShape.NonVisualConnectionShapeProperties?.NonVisualDrawingProperties?.Name?.Value,
                _ => null,
            };
            if (name?.StartsWith("Scribble", System.StringComparison.OrdinalIgnoreCase) == true)
            {
                return LineType.Scribble;
            }

            if (name?.StartsWith("Freeform", System.StringComparison.OrdinalIgnoreCase) == true)
            {
                return LineType.FreeformShape;
            }

            return LineType.Curve;
        }
    }

    public Point StartPoint
    {
        get
        {
            var transform = this.Transform;
            var flipH = transform.HorizontalFlip?.Value == true;
            var flipV = transform.VerticalFlip?.Value == true;
            var startX = flipH ? this.lineShape.X + this.lineShape.Width : this.lineShape.X;
            var startY = flipV ? this.lineShape.Y + this.lineShape.Height : this.lineShape.Y;
            return new Point(startX, startY);
        }
        set => this.UpdateEndpoints(value, this.EndPoint);
    }

    public Point EndPoint
    {
        get
        {
            var transform = this.Transform;
            var flipH = transform.HorizontalFlip?.Value == true;
            var flipV = transform.VerticalFlip?.Value == true;
            var endX = flipH ? this.lineShape.X : this.lineShape.X + this.lineShape.Width;
            var endY = flipV ? this.lineShape.Y : this.lineShape.Y + this.lineShape.Height;
            return new Point(endX, endY);
        }
        set => this.UpdateEndpoints(this.StartPoint, value);
    }

    public IReadOnlyList<Point> Points
    {
        get
        {
            var path = this.ShapeProperties.GetFirstChild<CustomGeometry>()?.PathList?.GetFirstChild<Path>();
            if (path?.Width?.Value is not { } pathWidth || path.Height?.Value is not { } pathHeight || pathWidth == 0 || pathHeight == 0)
            {
                return [this.StartPoint, this.EndPoint];
            }

            var points = path.Descendants<DocumentFormat.OpenXml.Drawing.Point>()
                .Where(point => point.X?.Value is not null && point.Y?.Value is not null)
                .Select(point => new Point(
                    decimal.Round(this.lineShape.X + decimal.Parse(point.X!.Value!, CultureInfo.InvariantCulture) / pathWidth * this.lineShape.Width, 6),
                    decimal.Round(this.lineShape.Y + decimal.Parse(point.Y!.Value!, CultureInfo.InvariantCulture) / pathHeight * this.lineShape.Height, 6)))
                .ToArray();

            return points.Length > 0 ? points : [this.StartPoint, this.EndPoint];
        }
    }

    private P.ShapeProperties ShapeProperties =>
        this.shapeElement.GetFirstChild<P.ShapeProperties>()
        ?? throw new SCException("Line shape does not contain shape properties.");

    private Transform2D Transform => this.ShapeProperties.Transform2D
        ?? throw new SCException("Line shape does not contain a 2D transform.");

    private void UpdateEndpoints(Point startPoint, Point endPoint)
    {
        var x = System.Math.Min(startPoint.X, endPoint.X);
        var y = System.Math.Min(startPoint.Y, endPoint.Y);
        var width = System.Math.Abs(endPoint.X - startPoint.X);
        var height = System.Math.Abs(endPoint.Y - startPoint.Y);
        var transform = this.Transform;

        transform.Offset ??= new Offset();
        transform.Extents ??= new Extents();
        transform.Offset.X = new Points(x).AsEmus();
        transform.Offset.Y = new Points(y).AsEmus();
        transform.Extents.Cx = new Points(width).AsEmus();
        transform.Extents.Cy = new Points(height).AsEmus();
        transform.HorizontalFlip = new BooleanValue(startPoint.X > endPoint.X);
        transform.VerticalFlip = new BooleanValue(startPoint.Y > endPoint.Y);
    }
}