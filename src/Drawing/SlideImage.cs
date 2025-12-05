using System;
using System.Collections.Generic;
using System.IO;
using ShapeCrawler.Slides;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Drawing;

/// <summary>
///     Renders a slide to an image.
/// </summary>
internal sealed class SlideImage(RemovedSlide slide)
{
    private const int EmusPerPixel = 9525; // EMUs to pixels conversion factor (96 DPI)
    private const float PointsToPixels = 96f / 72f; // Points to pixels conversion factor
    private const double Epsilon = 1e-6;
    private static readonly Dictionary<string, Func<A.ColorScheme, A.Color2Type?>> SchemeColorSelectors = new(StringComparer.Ordinal)
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

    private static SKColor ApplyShade(SKColor color, int shadeValue)
    {
        var shadeFactor = shadeValue / 100000f;

        return new SKColor(
            (byte)(color.Red * shadeFactor),
            (byte)(color.Green * shadeFactor),
            (byte)(color.Blue * shadeFactor),
            color.Alpha);
    }
    
    private static float GetStyleOutlineWidth(IShape shape)
    {
        var pShape = shape.SDKOpenXmlElement as P.Shape;
        if (pShape is null)
        {
            return 0;
        }

        var style = pShape.ShapeStyle;
        var lineRef = style?.LineReference;
        if (lineRef?.Index is null || lineRef.Index.Value == 0)
        {
            return 0;
        }

        // Default line width based on index (idx="2" typically means ~1.5pt line)
        // This is a simplification - proper implementation would look up theme line styles
        var defaultWidth = lineRef.Index.Value * 0.75f;
        return defaultWidth * PointsToPixels;
    }
    
    private static void ApplyRotation(SKCanvas canvas, IShape shape, float x, float y, float width, float height)
    {
        if (Math.Abs(shape.Rotation) > Epsilon)
        {
            var centerX = x + (width / 2);
            var centerY = y + (height / 2);
            canvas.RotateDegrees((float)shape.Rotation, centerX, centerY);
        }
    }
    
    private static float GetShapeOutlineWidth(IShape shape)
    {
        var shapeOutline = shape.Outline;

        // Check for explicit outline weight first
        if (shapeOutline is not null && shapeOutline.Weight > 0)
        {
            return (float)shapeOutline.Weight * PointsToPixels;
        }

        // Check for style-based outline width
        var styleWidth = GetStyleOutlineWidth(shape);
        return styleWidth;
    }
    
    private static SKColor ParseHexColor(string hex, double alphaPercentage)
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
    
    private SKColor GetSkColor()
    {
        var hex = slide.Fill.Color!.TrimStart('#');
        
        // Validate hex length before parsing
        if (hex.Length != 6 && hex.Length != 8)
        {
            return SKColors.White; // used by the PowerPoint application as the default background color
        }
        
        return ParseHexColor(hex, 100);
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
        ApplyRotation(canvas, shape, x, y, width, height);

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
        ApplyRotation(canvas, shape, x, y, width, height);

        this.RenderEllipseFill(canvas, shape, rect);
        this.RenderEllipseOutline(canvas, shape, rect);

        canvas.Restore();
    }

    private void RenderFill(SKCanvas canvas, IShape shape, SKRect rect, float cornerRadius)
    {
        var fillColor = this.GetShapeFillColor(shape);
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint
        {
            Color = fillColor.Value,
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

    private void RenderOutline(SKCanvas canvas, IShape shape, SKRect rect, float cornerRadius)
    {
        var outlineColor = this.GetShapeOutlineColor(shape);
        var strokeWidth = GetShapeOutlineWidth(shape);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint
        {
            Color = outlineColor.Value,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = strokeWidth,
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
        var fillColor = this.GetShapeFillColor(shape);
        if (fillColor is null)
        {
            return;
        }

        using var fillPaint = new SKPaint
        {
            Color = fillColor.Value,
            Style = SKPaintStyle.Fill,
            IsAntialias = true
        };

        canvas.DrawOval(rect, fillPaint);
    }

    private void RenderEllipseOutline(SKCanvas canvas, IShape shape, SKRect rect)
    {
        var outlineColor = this.GetShapeOutlineColor(shape);
        var strokeWidth = GetShapeOutlineWidth(shape);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint
        {
            Color = outlineColor.Value,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = strokeWidth,
            IsAntialias = true
        };

        canvas.DrawOval(rect, outlinePaint);
    }

    private SKColor? GetShapeOutlineColor(IShape shape)
    {
        var shapeOutline = shape.Outline;

        // Check for explicit outline color first
        if (shapeOutline?.HexColor is not null)
        {
            return ParseHexColor(shapeOutline.HexColor, 100);
        }

        // Check for style-based outline (lnRef with scheme color)
        var styleColor = this.GetStyleOutlineColor(shape);
        if (styleColor is not null)
        {
            return styleColor;
        }

        return null;
    }

    private SKColor? GetStyleOutlineColor(IShape shape)
    {
        if (shape.SDKOpenXmlElement is not P.Shape { ShapeStyle.LineReference: { } lineRef })
        {
            return null;
        }

        var schemeColor = lineRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor is null)
        {
            return null;
        }
        
        var schemeColorValue = schemeColor.Val?.InnerText;
        if (schemeColorValue is null)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);
        if (hexColor is null)
        {
            return null;
        }

        var baseColor = ParseHexColor(hexColor, 100);
        var shadeValue = schemeColor.GetFirstChild<A.Shade>()?.Val?.Value;

        return shadeValue is null
            ? baseColor
            : ApplyShade(baseColor, shadeValue.Value);
    }

    private SKColor? GetShapeFillColor(IShape shape)
    {
        var shapeFill = shape.Fill;

        // Check for explicit solid fill first
        if (shapeFill is { Type: FillType.Solid, Color: not null })
        {
            return ParseHexColor(shapeFill.Color, shapeFill.Alpha);
        }

        // Check for style-based fill (fillRef with scheme color)
        if (shapeFill is null || shapeFill.Type == FillType.NoFill)
        {
            var styleColor = this.GetStyleFillColor(shape);
            if (styleColor is not null)
            {
                return styleColor;
            }
        }

        return null;
    }

    private SKColor? GetStyleFillColor(IShape shape)
    {
        var pShape = shape.SDKOpenXmlElement as P.Shape;
        if (pShape is null)
        {
            return null;
        }

        var style = pShape.ShapeStyle;
        var fillRef = style?.FillReference;
        if (fillRef is null)
        {
            return null;
        }

        var schemeColor = fillRef.GetFirstChild<A.SchemeColor>();
        if (schemeColor?.Val is null)
        {
            return null;
        }

        var schemeColorValue = schemeColor.Val?.InnerText;
        if (schemeColorValue is null)
        {
            return null;
        }

        var hexColor = this.ResolveSchemeColor(schemeColorValue);

        return hexColor is not null ? ParseHexColor(hexColor, 100) : null;
    }

    private string? ResolveSchemeColor(string schemeColorName)
    {
        var colorScheme = this.GetColorScheme();
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

    private A.ColorScheme? GetColorScheme()
    {
        var layoutPart = slide.SlidePart.SlideLayoutPart;
        var masterPart = layoutPart?.SlideMasterPart;
        var themePart = masterPart?.ThemePart;
        var themeElements = themePart?.Theme.ThemeElements;
        
        return themeElements?.ColorScheme;
    }
}