using System;
using System.Collections.Generic;
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
            case Geometry.Rectangle:
            case Geometry.RoundedRectangle:
                this.RenderRectangle(canvas);
                break;
            case Geometry.Ellipse:
                this.RenderEllipse(canvas);
                break;
            case Geometry.Line:
            case Geometry.LineInverse:
            case Geometry.Triangle:
            case Geometry.RightTriangle:
            case Geometry.Diamond:
            case Geometry.Parallelogram:
            case Geometry.Trapezoid:
            case Geometry.NonIsoscelesTrapezoid:
            case Geometry.Pentagon:
            case Geometry.Hexagon:
            case Geometry.Heptagon:
            case Geometry.Octagon:
            case Geometry.Decagon:
            case Geometry.Dodecagon:
            case Geometry.Star4:
            case Geometry.Star5:
            case Geometry.Star6:
            case Geometry.Star7:
            case Geometry.Star8:
            case Geometry.Star10:
            case Geometry.Star12:
            case Geometry.Star16:
            case Geometry.Star24:
            case Geometry.Star32:
            case Geometry.SingleCornerRoundedRectangle:
            case Geometry.TopCornersRoundedRectangle:
            case Geometry.DiagonalCornersRoundedRectangle:
            case Geometry.SnipRoundRectangle:
            case Geometry.Snip1Rectangle:
            case Geometry.Snip2SameRectangle:
            case Geometry.Snip2DiagonalRectangle:
            case Geometry.Plaque:
            case Geometry.Teardrop:
            case Geometry.HomePlate:
            case Geometry.Chevron:
            case Geometry.PieWedge:
            case Geometry.Pie:
            case Geometry.BlockArc:
            case Geometry.Donut:
            case Geometry.NoSmoking:
            case Geometry.RightArrow:
            case Geometry.LeftArrow:
            case Geometry.UpArrow:
            case Geometry.DownArrow:
            case Geometry.StripedRightArrow:
            case Geometry.NotchedRightArrow:
            case Geometry.BentUpArrow:
            case Geometry.LeftRightArrow:
            case Geometry.UpDownArrow:
            case Geometry.LeftUpArrow:
            case Geometry.LeftRightUpArrow:
            case Geometry.QuadArrow:
            case Geometry.LeftArrowCallout:
            case Geometry.RightArrowCallout:
            case Geometry.UpArrowCallout:
            case Geometry.DownArrowCallout:
            case Geometry.LeftRightArrowCallout:
            case Geometry.UpDownArrowCallout:
            case Geometry.QuadArrowCallout:
            case Geometry.BentArrow:
            case Geometry.UTurnArrow:
            case Geometry.CircularArrow:
            case Geometry.LeftCircularArrow:
            case Geometry.LeftRightCircularArrow:
            case Geometry.CurvedRightArrow:
            case Geometry.CurvedLeftArrow:
            case Geometry.CurvedUpArrow:
            case Geometry.CurvedDownArrow:
            case Geometry.SwooshArrow:
            case Geometry.Cube:
            case Geometry.Can:
            case Geometry.LightningBolt:
            case Geometry.Heart:
            case Geometry.Sun:
            case Geometry.Moon:
            case Geometry.SmileyFace:
            case Geometry.IrregularSeal1:
            case Geometry.IrregularSeal2:
            case Geometry.FoldedCorner:
            case Geometry.Bevel:
            case Geometry.Frame:
            case Geometry.HalfFrame:
            case Geometry.Corner:
            case Geometry.DiagonalStripe:
            case Geometry.Chord:
            case Geometry.Arc:
            case Geometry.LeftBracket:
            case Geometry.RightBracket:
            case Geometry.LeftBrace:
            case Geometry.RightBrace:
            case Geometry.BracketPair:
            case Geometry.BracePair:
            case Geometry.StraightConnector1:
            case Geometry.BentConnector2:
            case Geometry.BentConnector3:
            case Geometry.BentConnector4:
            case Geometry.BentConnector5:
            case Geometry.CurvedConnector2:
            case Geometry.CurvedConnector3:
            case Geometry.CurvedConnector4:
            case Geometry.CurvedConnector5:
            case Geometry.Callout1:
            case Geometry.Callout2:
            case Geometry.Callout3:
            case Geometry.AccentCallout1:
            case Geometry.AccentCallout2:
            case Geometry.AccentCallout3:
            case Geometry.BorderCallout1:
            case Geometry.BorderCallout2:
            case Geometry.BorderCallout3:
            case Geometry.AccentBorderCallout1:
            case Geometry.AccentBorderCallout2:
            case Geometry.AccentBorderCallout3:
            case Geometry.WedgeRectangleCallout:
            case Geometry.WedgeRoundRectangleCallout:
            case Geometry.WedgeEllipseCallout:
            case Geometry.CloudCallout:
            case Geometry.Cloud:
            case Geometry.Ribbon:
            case Geometry.Ribbon2:
            case Geometry.EllipseRibbon:
            case Geometry.EllipseRibbon2:
            case Geometry.LeftRightRibbon:
            case Geometry.VerticalScroll:
            case Geometry.HorizontalScroll:
            case Geometry.Wave:
            case Geometry.DoubleWave:
            case Geometry.Plus:
            case Geometry.FlowChartProcess:
            case Geometry.FlowChartDecision:
            case Geometry.FlowChartInputOutput:
            case Geometry.FlowChartPredefinedProcess:
            case Geometry.FlowChartInternalStorage:
            case Geometry.FlowChartDocument:
            case Geometry.FlowChartMultidocument:
            case Geometry.FlowChartTerminator:
            case Geometry.FlowChartPreparation:
            case Geometry.FlowChartManualInput:
            case Geometry.FlowChartManualOperation:
            case Geometry.FlowChartConnector:
            case Geometry.FlowChartPunchedCard:
            case Geometry.FlowChartPunchedTape:
            case Geometry.FlowChartSummingJunction:
            case Geometry.FlowChartOr:
            case Geometry.FlowChartCollate:
            case Geometry.FlowChartSort:
            case Geometry.FlowChartExtract:
            case Geometry.FlowChartMerge:
            case Geometry.FlowChartOfflineStorage:
            case Geometry.FlowChartOnlineStorage:
            case Geometry.FlowChartMagneticTape:
            case Geometry.FlowChartMagneticDisk:
            case Geometry.FlowChartMagneticDrum:
            case Geometry.FlowChartDisplay:
            case Geometry.FlowChartDelay:
            case Geometry.FlowChartAlternateProcess:
            case Geometry.FlowChartOffpageConnector:
            case Geometry.ActionButtonBlank:
            case Geometry.ActionButtonHome:
            case Geometry.ActionButtonHelp:
            case Geometry.ActionButtonInformation:
            case Geometry.ActionButtonForwardNext:
            case Geometry.ActionButtonBackPrevious:
            case Geometry.ActionButtonEnd:
            case Geometry.ActionButtonBeginning:
            case Geometry.ActionButtonReturn:
            case Geometry.ActionButtonDocument:
            case Geometry.ActionButtonSound:
            case Geometry.ActionButtonMovie:
            case Geometry.Gear6:
            case Geometry.Gear9:
            case Geometry.Funnel:
            case Geometry.MathPlus:
            case Geometry.MathMinus:
            case Geometry.MathMultiply:
            case Geometry.MathDivide:
            case Geometry.MathEqual:
            case Geometry.MathNotEqual:
            case Geometry.CornerTabs:
            case Geometry.SquareTabs:
            case Geometry.PlaqueTabs:
            case Geometry.ChartX:
            case Geometry.ChartStar:
            case Geometry.ChartPlus:
            case Geometry.Custom:
                break;
            default:
                throw new ArgumentOutOfRangeException();
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

    private decimal GetShapeOutlineWidth()
    {
        var shapeOutline = this.Outline;

        if (shapeOutline is not null && shapeOutline.Weight > 0)
        {
            return new Points(shapeOutline.Weight).AsPixels();
        }

        var styleWidth = GetStyleOutlineWidth();
        return styleWidth;
    }

    private decimal GetStyleOutlineWidth()
    {
        if (this.SDKOpenXmlElement is not P.Shape pShape)
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
        if (this.PShapeTreeElement is not P.Shape { ShapeStyle.LineReference: { } lineRef })
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