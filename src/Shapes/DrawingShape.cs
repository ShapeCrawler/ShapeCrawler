using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Units;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using Position = ShapeCrawler.Positions.Position;

namespace ShapeCrawler.Shapes;

internal class DrawingShape(Position position, ShapeSize shapeSize, ShapeId shapeId, OpenXmlElement pShapeTreeElement)
    : Shape(position, shapeSize, shapeId, pShapeTreeElement)
{
    private const double Epsilon = 1e-6;
    private static readonly Dictionary<string, Func<A.ColorScheme, A.Color2Type?>> SchemeColorSelectors =
        new(StringComparer.Ordinal)
        {
            { "dk1", scheme => scheme.Dark1Color },
            { "lt1", scheme => scheme.Light1Color },
            { "dk2", scheme => scheme.Dark2Color },
            { "lt2", scheme => scheme.Light2Color },
            { "accent1", scheme => scheme.Accent1Color },
            { "accent2", scheme => scheme.Accent2Color },
            { "accent3", scheme => scheme.Accent3Color },
            { "accent4", scheme => scheme.Accent4Color },
            { "accent5", scheme => scheme.Accent5Color },
            { "accent6", scheme => scheme.Accent6Color },
            { "hlink", scheme => scheme.Hyperlink },
            { "folHlink", scheme => scheme.FollowedHyperlinkColor }
        };
    
    /// <summary>
    ///     Renders the current shape onto the provided canvas.
    /// </summary>
    /// <param name="canvas">Target canvas.</param>
    internal virtual void Render(SKCanvas canvas)
    {
        switch (this.GeometryType)
        {
            case Geometry.Line:
            case Geometry.LineInverse:
                this.RenderLine(canvas);
                break;
            case Geometry.Rectangle:
            case Geometry.RoundedRectangle:
                this.RenderRectangle(canvas);
                break;
            case Geometry.Ellipse:
                this.RenderEllipse(canvas);
                break;
            default:
                throw new SCException("Unsupported shape geometry type.");
        }
    }

    private static string? GetHexFromColorElement(A.Color2Type colorElement)
    {
        var rgbColor = colorElement.RgbColorModelHex;
        if (rgbColor?.Val?.Value is { } rgb)
        {
            return rgb;
        }

        var sysColor = colorElement.SystemColor;
        return sysColor?.LastColor?.Value;
    }

    private static SKColor ApplyShade(SKColor color, int shadeValue)
    {
        var shadeFactor = shadeValue / 100_000f;

        return new SKColor(
            (byte)(color.Red * shadeFactor),
            (byte)(color.Green * shadeFactor),
            (byte)(color.Blue * shadeFactor),
            color.Alpha);
    }
    
    private static SKColor ApplyShadeIfNeeded(A.SchemeColor schemeColor, string hexColor)
    {
        var baseColor = new Color(hexColor).AsSkColor();
        var shadeValue = schemeColor.GetFirstChild<A.Shade>()?.Val?.Value;

        return shadeValue is null
            ? baseColor
            : ApplyShade(baseColor, shadeValue.Value);
    }

    private decimal GetShapeOutlineWidth()
    {
        var shapeOutline = this.Outline;

        if (shapeOutline.Weight > 0)
        {
            return new Points(shapeOutline.Weight).AsPixels();
        }

        var styleWidth = GetStyleOutlineWidth();
        return styleWidth;
    }

    private decimal GetStyleOutlineWidth()
    {
        var lineRef = this.PShapeTreeElement switch
        {
            P.Shape pShape => pShape.ShapeStyle?.LineReference,
            P.ConnectionShape pConnectionShape => pConnectionShape.ShapeStyle?.LineReference,
            _ => null
        };

        if (lineRef is null)
        {
            return 0;
        }

        if (lineRef.Index is null || lineRef.Index.Value == 0)
        {
            return 0;
        }

        var defaultWidth = lineRef.Index.Value * 0.75m;

        return new Points(defaultWidth).AsPixels();
    }

    private void ApplyRotation(SKCanvas canvas)
    {
        if (Math.Abs(this.Rotation) > Epsilon)
        {
            var centerX = this.X + (this.Width / 2);
            var centerY = this.Y + (this.Height / 2);
            canvas.RotateDegrees(
                (float)this.Rotation,
                (float)new Points(centerX).AsPixels(),
                (float)new Points(centerY).AsPixels()
            );
        }
    }

    private void RenderRectangle(SKCanvas canvas)
    {
        var x = new Points(this.X).AsPixels();
        var y = new Points(this.Y).AsPixels();
        var width = new Points(this.Width).AsPixels();
        var height = new Points(this.Height).AsPixels();
        var rect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        var cornerRadius = 0m;
        if (this.GeometryType == Geometry.RoundedRectangle)
        {
            var shortestSide = Math.Min(width, height);
            cornerRadius = this.CornerSize / 100m * (shortestSide / 2m);
        }

        canvas.Save();
        ApplyRotation(canvas);

        this.RenderFill(canvas, rect, cornerRadius);
        this.RenderOutline(canvas, rect, cornerRadius);

        canvas.Restore();
    }

    private void RenderEllipse(SKCanvas canvas)
    {
        var x = new Points(this.X).AsPixels();
        var y = new Points(this.Y).AsPixels();
        var width = new Points(this.Width).AsPixels();
        var height = new Points(this.Height).AsPixels();
        var rect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        canvas.Save();
        ApplyRotation(canvas);

        this.RenderEllipseFill(canvas, rect);
        this.RenderEllipseOutline(canvas, rect);

        canvas.Restore();
    }

    private void RenderLine(SKCanvas canvas)
    {
        var outlineColor = this.GetShapeOutlineColor();
        var strokeWidth = this.GetShapeOutlineWidth();
        var line = this.Line;
        if (outlineColor is null || strokeWidth <= 0 || line is null)
        {
            return;
        }

        var startPoint = new SKPoint(
            (float)new Points(line.StartPoint.X).AsPixels(),
            (float)new Points(line.StartPoint.Y).AsPixels());
        var endPoint = new SKPoint(
            (float)new Points(line.EndPoint.X).AsPixels(),
            (float)new Points(line.EndPoint.Y).AsPixels());

        using var outlinePaint = new SKPaint
        {
            Color = outlineColor.Value,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = (float)strokeWidth,
            IsAntialias = true
        };

        canvas.Save();
        this.ApplyRotation(canvas);
        canvas.DrawLine(startPoint, endPoint, outlinePaint);

        this.RenderArrows(canvas, startPoint, endPoint, outlinePaint);

        canvas.Restore();
    }

    private void RenderArrows(SKCanvas canvas, SKPoint startPoint, SKPoint endPoint, SKPaint outlinePaint)
    {
        var pShapeProperties = this.PShapeTreeElement.Descendants<P.ShapeProperties>().FirstOrDefault();
        var aOutline = pShapeProperties?.GetFirstChild<A.Outline>();
        if (aOutline == null)
        {
            return;
        }

        var headEnd = aOutline.GetFirstChild<A.HeadEnd>();
        this.RenderArrowEnd(canvas, endPoint, startPoint, headEnd?.Type, outlinePaint);

        var tailEnd = aOutline.GetFirstChild<A.TailEnd>();
        this.RenderArrowEnd(canvas, startPoint, endPoint, tailEnd?.Type, outlinePaint);
    }

    private void RenderArrowEnd(SKCanvas canvas, SKPoint tail, SKPoint tip, EnumValue<A.LineEndValues>? type, SKPaint paint)
    {
        if (type?.Value is not null && type.Value != A.LineEndValues.None)
        {
            this.RenderArrowHead(canvas, tail, tip, type.Value, paint);
        }
    }

    private void RenderArrowHead(SKCanvas canvas, SKPoint tail, SKPoint tip, A.LineEndValues type, SKPaint linePaint)
    {
        var angle = Math.Atan2(tip.Y - tail.Y, tip.X - tail.X);
        var arrowSize = linePaint.StrokeWidth * 3;

        using var paint = new SKPaint();
        paint.Color = linePaint.Color;
        paint.IsAntialias = true;

        var path = new SKPath();

        if (type == A.LineEndValues.Triangle)
        {
            paint.Style = SKPaintStyle.Fill;
            path.MoveTo(0, 0);
            path.LineTo(-arrowSize, -arrowSize / 2);
            path.LineTo(-arrowSize, arrowSize / 2);
            path.Close();
        }
        else if (type == A.LineEndValues.Stealth)
        {
            paint.Style = SKPaintStyle.Fill;
            path.MoveTo(0, 0);
            path.LineTo(-arrowSize, -arrowSize / 2);
            path.LineTo(-arrowSize * 0.6f, 0);
            path.LineTo(-arrowSize, arrowSize / 2);
            path.Close();
        }
        else if (type == A.LineEndValues.Arrow)
        {
            paint.Style = SKPaintStyle.Stroke;
            paint.StrokeWidth = linePaint.StrokeWidth;
            path.MoveTo(-arrowSize, -arrowSize / 2);
            path.LineTo(0, 0);
            path.LineTo(-arrowSize, arrowSize / 2);
        }
        else
        {
            // Default to Triangle for other types for now
            paint.Style = SKPaintStyle.Fill;
            path.MoveTo(0, 0);
            path.LineTo(-arrowSize, -arrowSize / 2);
            path.LineTo(-arrowSize, arrowSize / 2);
            path.Close();
        }

        canvas.Save();
        canvas.Translate(tip.X, tip.Y);
        canvas.RotateDegrees((float)(angle * 180 / Math.PI));
        canvas.DrawPath(path, paint);
        canvas.Restore();
    }

    private void RenderFill(SKCanvas canvas, SKRect rect, decimal cornerRadius)
    {
        var fillColor = this.GetShapeFillColor();
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint();
        fillPaint.Color = fillColor.Value;
        fillPaint.Style = SKPaintStyle.Fill;
        fillPaint.IsAntialias = true;

        if (cornerRadius > 0)
        {
            canvas.DrawRoundRect(rect, (float)cornerRadius, (float)cornerRadius, fillPaint);
        }
        else
        {
            canvas.DrawRect(rect, fillPaint);
        }
    }

    private void RenderOutline(SKCanvas canvas, SKRect rect, decimal cornerRadius)
    {
        var outlineColor = this.GetShapeOutlineColor();
        var strokeWidth = GetShapeOutlineWidth();

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = outlineColor.Value;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = (float)strokeWidth;
        outlinePaint.IsAntialias = true;

        if (cornerRadius > 0)
        {
            canvas.DrawRoundRect(rect, (float)cornerRadius, (float)cornerRadius, outlinePaint);
        }
        else
        {
            canvas.DrawRect(rect, outlinePaint);
        }
    }

    private void RenderEllipseFill(SKCanvas canvas, SKRect rect)
    {
        var fillColor = this.GetShapeFillColor();
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint();
        fillPaint.Color = fillColor.Value;
        fillPaint.Style = SKPaintStyle.Fill;
        fillPaint.IsAntialias = true;

        canvas.DrawOval(rect, fillPaint);
    }

    private void RenderEllipseOutline(SKCanvas canvas, SKRect rect)
    {
        var outlineColor = this.GetShapeOutlineColor();
        var strokeWidth = GetShapeOutlineWidth();

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint();
        outlinePaint.Color = outlineColor.Value;
        outlinePaint.Style = SKPaintStyle.Stroke;
        outlinePaint.StrokeWidth = (float)strokeWidth;
        outlinePaint.IsAntialias = true;

        canvas.DrawOval(rect, outlinePaint);
    }

    private SKColor? GetShapeFillColor()
    {
        var shapeFill = this.Fill;

        if (shapeFill is { Type: FillType.Solid, Color: not null })
        {
            return new Color(shapeFill.Color).AsSkColor();
        }

        if (shapeFill.Type != FillType.NoFill)
        {
            return null;
        }

        var styleColor = this.GetStyleFillColor();
        return styleColor;
    }

    private SKColor? GetShapeOutlineColor()
    {
        var shapeOutline = this.Outline;

        if (shapeOutline?.HexColor is not null)
        {
            return new Color(shapeOutline.HexColor).AsSkColor();
        }

        var styleColor = this.GetStyleOutlineColor();
        if (styleColor is not null)
        {
            return styleColor;
        }

        return null;
    }

    private SKColor? GetStyleOutlineColor()
    {
        var lineRef = this.GetLineReference();
        if (lineRef is null)
        {
            return null;
        }

        var schemeColor = lineRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor?.Val?.InnerText is not { } schemeColorValue)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);
        if (hexColor is null)
        {
            return null;
        }

        return ApplyShadeIfNeeded(schemeColor, hexColor);
    }

    private A.LineReference? GetLineReference()
    {
        return this.PShapeTreeElement switch
        {
            P.Shape pShape => pShape.ShapeStyle?.LineReference,
            P.ConnectionShape pConnectionShape => pConnectionShape.ShapeStyle?.LineReference,
            _ => null
        };
    }

    private SKColor? GetStyleFillColor()
    {
        if (this.PShapeTreeElement is not P.Shape pShape)
        {
            return null;
        }

        var fillRef = pShape.ShapeStyle?.FillReference;
        if (fillRef is null)
        {
            return null;
        }

        var schemeColor = fillRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor?.Val?.InnerText is not { } schemeColorValue)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);

        return hexColor is not null ? new Color(hexColor).AsSkColor() : null;
    }

    private string? ResolveSchemeColor(string schemeColorName)
    {
        var shapeColorScheme = new ShapeColorScheme(this.PShapeTreeElement);
        var colorScheme = shapeColorScheme.GetColorScheme();
        if (colorScheme is null)
        {
            return null;
        }

        if (!SchemeColorSelectors.TryGetValue(schemeColorName, out var selector))
        {
            return null;
        }

        var colorElement = selector(colorScheme);

        return colorElement is null ? null : GetHexFromColorElement(colorElement);
    }
}