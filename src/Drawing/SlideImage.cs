using System;
using System.IO;
using SkiaSharp;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Renders a slide to an image.
/// </summary>
internal sealed class SlideImage(ISlide slide)
{
    private const int EmusPerPixel = 9525; // EMUs to pixels conversion factor (96 DPI)
    private const float PointsToPixels = 96f / 72f; // Points to pixels conversion factor

    /// <summary>
    ///     Saves the slide to the specified stream in the given image format.
    /// </summary>
    internal void Save(Stream stream, SKEncodedImageFormat format)
    {
        var presPart = slide.GetSDKPresentationPart();
        var pSlideSize = presPart.Presentation.SlideSize!;
        var width = pSlideSize.Cx!.Value / EmusPerPixel;
        var height = pSlideSize.Cy!.Value / EmusPerPixel;

        using var surface = SKSurface.Create(new SKImageInfo(width, height));
        var canvas = surface.Canvas;

        this.RenderBackground(canvas);
        this.RenderShapes(canvas);

        using var image = surface.Snapshot();
        using var data = image.Encode(format, 100);
        data.SaveTo(stream);
    }

    private SKColor GetSkColor()
    {
        var hex = slide.Fill.Color!.TrimStart('#');

        if (hex.Length == 6)
        {
            // Parse RGB (RRGGBB)
            var r = Convert.ToByte(hex[..2], 16);
            var g = Convert.ToByte(hex.Substring(2, 2), 16);
            var b = Convert.ToByte(hex.Substring(4, 2), 16);
            return new SKColor(r, g, b);
        }

        if (hex.Length == 8)
        {
            // Parse ARGB (AARRGGBB)
            var a = Convert.ToByte(hex[..2], 16);
            var r = Convert.ToByte(hex.Substring(2, 2), 16);
            var g = Convert.ToByte(hex.Substring(4, 2), 16);
            var b = Convert.ToByte(hex.Substring(6, 2), 16);
            return new SKColor(r, g, b, a);
        }

        // Default to white if hex is invalid
        return SKColors.White;
    }

    private void RenderBackground(SKCanvas canvas)
    {
        var fill = slide.Fill;

        if (fill is { Type: FillType.Solid, Color: not null })
        {
            var skColor = this.GetSkColor();
            canvas.Clear(skColor);
        }
        else if (fill is { Type: FillType.Picture, Picture: not null })
        {
            var bytes = fill.Picture.AsByteArray();
            using var stream = new MemoryStream(bytes);
            using var bitmap = SKBitmap.Decode(stream);
            var destRect = new SKRect(0, 0, canvas.DeviceClipBounds.Width, canvas.DeviceClipBounds.Height);
            canvas.DrawBitmap(bitmap, destRect);
        }
        else
        {
            // Default to white for unsupported backgrounds
            canvas.Clear(SKColors.White);
        }
    }

    private void RenderShapes(SKCanvas canvas)
    {
        foreach (var shape in slide.Shapes)
        {
            if (shape.Hidden)
            {
                continue;
            }

            this.RenderShape(canvas, shape);
        }
    }

    private void RenderShape(SKCanvas canvas, IShape shape)
    {
        var geometryType = shape.GeometryType;

        switch (geometryType)
        {
            case Geometry.Rectangle:
            case Geometry.RoundedRectangle:
                this.RenderRectangle(canvas, shape);
                break;
            case Geometry.Ellipse:
                this.RenderEllipse(canvas, shape);
                break;
            default:
                // Other shapes not yet supported
                break;
        }
    }

    private void RenderRectangle(SKCanvas canvas, IShape shape)
    {
        var x = (float)shape.X * PointsToPixels;
        var y = (float)shape.Y * PointsToPixels;
        var width = (float)shape.Width * PointsToPixels;
        var height = (float)shape.Height * PointsToPixels;
        var rect = new SKRect(x, y, x + width, y + height);

        var cornerRadius = 0f;
        if (shape.GeometryType == Geometry.RoundedRectangle)
        {
            // CornerSize is percentage (0-100), where 100 = half of shortest side
            var shortestSide = Math.Min(width, height);
            cornerRadius = (float)shape.CornerSize / 100f * (shortestSide / 2f);
        }

        canvas.Save();
        this.ApplyRotation(canvas, shape, x, y, width, height);

        this.RenderFill(canvas, shape, rect, cornerRadius);
        this.RenderOutline(canvas, shape, rect, cornerRadius);

        canvas.Restore();
    }

    private void RenderEllipse(SKCanvas canvas, IShape shape)
    {
        var x = (float)shape.X * PointsToPixels;
        var y = (float)shape.Y * PointsToPixels;
        var width = (float)shape.Width * PointsToPixels;
        var height = (float)shape.Height * PointsToPixels;
        var rect = new SKRect(x, y, x + width, y + height);

        canvas.Save();
        this.ApplyRotation(canvas, shape, x, y, width, height);

        this.RenderEllipseFill(canvas, shape, rect);
        this.RenderEllipseOutline(canvas, shape, rect);

        canvas.Restore();
    }

    private void ApplyRotation(SKCanvas canvas, IShape shape, float x, float y, float width, float height)
    {
        if (shape.Rotation != 0)
        {
            var centerX = x + (width / 2);
            var centerY = y + (height / 2);
            canvas.RotateDegrees((float)shape.Rotation, centerX, centerY);
        }
    }

    private void RenderFill(SKCanvas canvas, IShape shape, SKRect rect, float cornerRadius)
    {
        var shapeFill = shape.Fill;
        if (shapeFill is null || shapeFill.Type == FillType.NoFill)
        {
            return;
        }

        if (shapeFill.Type == FillType.Solid && shapeFill.Color is not null)
        {
            var fillColor = this.ParseHexColor(shapeFill.Color, shapeFill.Alpha);
            using var fillPaint = new SKPaint
            {
                Color = fillColor,
                Style = SKPaintStyle.Fill,
                IsAntialias = true
            };

            if (cornerRadius > 0)
            {
                canvas.DrawRoundRect(rect, cornerRadius, cornerRadius, fillPaint);
            }
            else
            {
                canvas.DrawRect(rect, fillPaint);
            }
        }
    }

    private void RenderOutline(SKCanvas canvas, IShape shape, SKRect rect, float cornerRadius)
    {
        var shapeOutline = shape.Outline;
        if (shapeOutline is null || shapeOutline.Weight <= 0)
        {
            return;
        }

        var outlineColor = shapeOutline.HexColor is not null
            ? this.ParseHexColor(shapeOutline.HexColor, 100)
            : SKColors.Black;

        using var outlinePaint = new SKPaint
        {
            Color = outlineColor,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = (float)shapeOutline.Weight * PointsToPixels,
            IsAntialias = true
        };

        if (cornerRadius > 0)
        {
            canvas.DrawRoundRect(rect, cornerRadius, cornerRadius, outlinePaint);
        }
        else
        {
            canvas.DrawRect(rect, outlinePaint);
        }
    }

    private void RenderEllipseFill(SKCanvas canvas, IShape shape, SKRect rect)
    {
        var shapeFill = shape.Fill;
        if (shapeFill is null || shapeFill.Type == FillType.NoFill)
        {
            return;
        }

        if (shapeFill.Type == FillType.Solid && shapeFill.Color is not null)
        {
            var fillColor = this.ParseHexColor(shapeFill.Color, shapeFill.Alpha);
            using var fillPaint = new SKPaint
            {
                Color = fillColor,
                Style = SKPaintStyle.Fill,
                IsAntialias = true
            };

            canvas.DrawOval(rect, fillPaint);
        }
    }

    private void RenderEllipseOutline(SKCanvas canvas, IShape shape, SKRect rect)
    {
        var shapeOutline = shape.Outline;
        if (shapeOutline is null || shapeOutline.Weight <= 0)
        {
            return;
        }

        var outlineColor = shapeOutline.HexColor is not null
            ? this.ParseHexColor(shapeOutline.HexColor, 100)
            : SKColors.Black;

        using var outlinePaint = new SKPaint
        {
            Color = outlineColor,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = (float)shapeOutline.Weight * PointsToPixels,
            IsAntialias = true
        };

        canvas.DrawOval(rect, outlinePaint);
    }

    private SKColor ParseHexColor(string hex, double alphaPercentage)
    {
        hex = hex.TrimStart('#');

        byte r, g, b;
        byte a = (byte)(alphaPercentage / 100.0 * 255);

        if (hex.Length == 6)
        {
            r = Convert.ToByte(hex[..2], 16);
            g = Convert.ToByte(hex.Substring(2, 2), 16);
            b = Convert.ToByte(hex.Substring(4, 2), 16);
        }
        else if (hex.Length == 8)
        {
            a = Convert.ToByte(hex[..2], 16);
            r = Convert.ToByte(hex.Substring(2, 2), 16);
            g = Convert.ToByte(hex.Substring(4, 2), 16);
            b = Convert.ToByte(hex.Substring(6, 2), 16);
        }
        else
        {
            return SKColors.Transparent;
        }

        return new SKColor(r, g, b, a);
    }
}