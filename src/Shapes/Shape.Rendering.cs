using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using ShapeCrawler.Units;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal partial class Shape
{
    private const double Epsilon = 1e-6;
    private static readonly TextDrawing TextDrawing = new();
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
        if (this.Picture is not null)
        {
            this.RenderPicture(canvas);
            return;
        }

        switch (this.GeometryType)
        {
            case Geometry.Rectangle:
            case Geometry.RoundedRectangle:
                this.RenderRectangle(canvas);
                break;
            case Geometry.Ellipse:
                this.RenderEllipse(canvas);
                break;
            default:
                this.RenderText(canvas);
                return;
        }

        this.RenderText(canvas);
    }

    private static void ApplyRotation(
        SKCanvas canvas,
        IShape shape,
        decimal x,
        decimal y,
        decimal width,
        decimal height)
    {
        if (Math.Abs(shape.Rotation) > Epsilon)
        {
            var centerX = x + (width / 2);
            var centerY = y + (height / 2);
            canvas.RotateDegrees(
                (float)shape.Rotation,
                (float)new Points(centerX).AsPixels(),
                (float)new Points(centerY).AsPixels()
            );
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

    private static decimal GetShapeOutlineWidth(IShape shape)
    {
        var shapeOutline = shape.Outline;

        if (shapeOutline is not null && shapeOutline.Weight > 0)
        {
            return new Points(shapeOutline.Weight).AsPixels();
        }

        var styleWidth = GetStyleOutlineWidth(shape);
        return styleWidth;
    }

    private static decimal GetStyleOutlineWidth(IShape shape)
    {
        if (shape.SDKOpenXmlElement is not P.Shape pShape)
        {
            return 0;
        }

        var lineRef = pShape.ShapeStyle?.LineReference;
        if (lineRef?.Index is null || lineRef.Index.Value == 0)
        {
            return 0;
        }

        var defaultWidth = lineRef.Index.Value * 0.75m;

        return new Points(defaultWidth).AsPixels();
    }

    private void RenderPicture(SKCanvas canvas)
    {
        var picture = this.Picture!;
        var image = picture.Image;
        if (image is null)
        {
            return;
        }

        var imageBytes = image.AsByteArray();
        using var bitmap = SKBitmap.Decode(imageBytes);
        if (bitmap is null)
        {
            return;
        }

        var x = new Points(this.X).AsPixels();
        var y = new Points(this.Y).AsPixels();
        var width = new Points(this.Width).AsPixels();
        var height = new Points(this.Height).AsPixels();

        canvas.Save();
        ApplyRotation(canvas, this, this.X, this.Y, this.Width, this.Height);

        var crop = picture.Crop;
        var srcLeft = (float)(bitmap.Width * (double)(crop.Left / 100m));
        var srcTop = (float)(bitmap.Height * (double)(crop.Top / 100m));
        var srcRight = (float)(bitmap.Width * (1 - (double)(crop.Right / 100m)));
        var srcBottom = (float)(bitmap.Height * (1 - (double)(crop.Bottom / 100m)));
        var srcRect = new SKRect(srcLeft, srcTop, srcRight, srcBottom);

        var destRect = new SKRect((float)x, (float)y, (float)(x + width), (float)(y + height));

        using var paint = new SKPaint { IsAntialias = true };

        var transparency = picture.Transparency;
        if (transparency > 0)
        {
            var alpha = (byte)(255 * (1 - (double)(transparency / 100m)));
            paint.Color = paint.Color.WithAlpha(alpha);
        }

        canvas.DrawBitmap(bitmap, srcRect, destRect, paint);
        canvas.Restore();
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
        ApplyRotation(canvas, this, this.X, this.Y, this.Width, this.Height);

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
        ApplyRotation(canvas, this, this.X, this.Y, this.Width, this.Height);

        this.RenderEllipseFill(canvas, rect);
        this.RenderEllipseOutline(canvas, rect);

        canvas.Restore();
    }

    private void RenderText(SKCanvas canvas)
    {
        if (this.TextBox is null)
        {
            return;
        }

        canvas.Save();
        ApplyRotation(canvas, this, this.X, this.Y, this.Width, this.Height);
        TextDrawing.Render(canvas, this);
        canvas.Restore();
    }

    private void RenderFill(SKCanvas canvas, SKRect rect, decimal cornerRadius)
    {
        var fillColor = this.GetShapeFillColor();
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
        var strokeWidth = GetShapeOutlineWidth(this);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint
        {
            Color = outlineColor.Value,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = (float)strokeWidth,
            IsAntialias = true
        };

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

        using var fillPaint = new SKPaint
        {
            Color = fillColor.Value,
            Style = SKPaintStyle.Fill,
            IsAntialias = true
        };

        canvas.DrawOval(rect, fillPaint);
    }

    private void RenderEllipseOutline(SKCanvas canvas, SKRect rect)
    {
        var outlineColor = this.GetShapeOutlineColor();
        var strokeWidth = GetShapeOutlineWidth(this);

        if (outlineColor is null || strokeWidth <= 0)
        {
            return;
        }

        using var outlinePaint = new SKPaint
        {
            Color = outlineColor.Value,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = (float)strokeWidth,
            IsAntialias = true
        };

        canvas.DrawOval(rect, outlinePaint);
    }

    private SKColor? GetShapeFillColor()
    {
        var shapeFill = this.Fill;

        if (shapeFill is { Type: FillType.Solid, Color: not null })
        {
            return new Color(shapeFill.Color).AsSkColor();
        }

        if (shapeFill is null || shapeFill.Type == FillType.NoFill)
        {
            var styleColor = this.GetStyleFillColor();
            if (styleColor is not null)
            {
                return styleColor;
            }
        }

        return null;
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
        if (pShapeTreeElement is not P.Shape { ShapeStyle.LineReference: { } lineRef })
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

        var baseColor = new Color(hexColor).AsSkColor();
        var shadeValue = schemeColor.GetFirstChild<A.Shade>()?.Val?.Value;

        return shadeValue is null
            ? baseColor
            : ApplyShade(baseColor, shadeValue.Value);
    }

    private SKColor? GetStyleFillColor()
    {
        if (pShapeTreeElement is not P.Shape pShape)
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
        var parentPart = new SCOpenXmlElement(pShapeTreeElement).ParentOpenXmlPart;

        return parentPart switch
        {
            SlidePart slidePart => this.GetColorSchemeFromSlidePart(slidePart),
            SlideLayoutPart slideLayoutPart => this.GetColorSchemeFromSlideLayoutPart(slideLayoutPart),
            SlideMasterPart slideMasterPart => this.GetColorSchemeFromSlideMasterPart(slideMasterPart),
            NotesSlidePart notesSlidePart => this.GetColorSchemeFromNotesSlidePart(notesSlidePart),
            _ => null
        };
    }

    private A.ColorScheme? GetColorSchemeFromSlidePart(SlidePart slidePart)
    {
        var slideLayoutPart = slidePart.SlideLayoutPart;
        if (slideLayoutPart is null)
        {
            return null;
        }

        return this.GetColorSchemeFromSlideLayoutPart(slideLayoutPart);
    }

    private A.ColorScheme? GetColorSchemeFromSlideLayoutPart(SlideLayoutPart slideLayoutPart)
    {
        var slideMasterPart = slideLayoutPart.SlideMasterPart;
        if (slideMasterPart is null)
        {
            return null;
        }

        return this.GetColorSchemeFromSlideMasterPart(slideMasterPart);
    }

    private A.ColorScheme? GetColorSchemeFromSlideMasterPart(SlideMasterPart slideMasterPart)
    {
        var themePart = slideMasterPart.ThemePart;
        var themeElements = themePart?.Theme.ThemeElements;

        return themeElements?.ColorScheme;
    }

    private A.ColorScheme? GetColorSchemeFromNotesSlidePart(NotesSlidePart notesSlidePart)
    {
        var notesMasterPart = notesSlidePart.NotesMasterPart;
        if (notesMasterPart is null)
        {
            return null;
        }

        var themePart = notesMasterPart.ThemePart;
        var themeElements = themePart?.Theme.ThemeElements;

        return themeElements?.ColorScheme;
    }
}
