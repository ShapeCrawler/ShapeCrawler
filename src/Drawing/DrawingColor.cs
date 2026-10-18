using System;
using DocumentFormat.OpenXml;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Drawing;

internal sealed class DrawingColor(SKColor color)
{
    private double red = color.Red / 255d;
    private double green = color.Green / 255d;
    private double blue = color.Blue / 255d;
    private double alpha = color.Alpha / 255d;

    internal void ApplyTransforms(OpenXmlElement? element)
    {
        if (element is null)
        {
            return;
        }

        foreach (var transform in element.ChildElements)
        {
            switch (transform)
            {
                case A.Tint { Val: not null } tint:
                    var tintFactor = tint.Val.Value / 100000d;
                    this.TransformRgb(tintFactor, 1 - tintFactor);
                    break;
                case A.Shade { Val: not null } shade:
                    this.TransformRgb(shade.Val.Value / 100000d, 0);
                    break;
                case A.LuminanceModulation { Val: not null } modulation:
                    this.TransformLuminance(modulation.Val.Value / 100000d, 0);
                    break;
                case A.LuminanceOffset { Val: not null } offset:
                    this.TransformLuminance(1, offset.Val.Value / 100000d);
                    break;
                case A.Alpha { Val: not null } opacity:
                    this.alpha = Math.Clamp(opacity.Val.Value / 100000d, 0, 1);
                    break;
                case A.AlphaModulation { Val: not null } opacityModulation:
                    this.alpha = Math.Clamp(this.alpha * opacityModulation.Val.Value / 100000d, 0, 1);
                    break;
                case A.AlphaOffset { Val: not null } opacityOffset:
                    this.alpha = Math.Clamp(this.alpha + opacityOffset.Val.Value / 100000d, 0, 1);
                    break;
            }
        }
    }

    internal SKColor AsSkColor() => new(ToByte(this.red), ToByte(this.green), ToByte(this.blue), ToByte(this.alpha));

    private static byte ToByte(double value) =>
        (byte)Math.Round(Math.Clamp(value, 0, 1) * 255, MidpointRounding.AwayFromZero);

    private static double ToLinear(double value) =>
        value <= 0.04045 ? value / 12.92 : Math.Pow((value + 0.055) / 1.055, 2.4);

    private static double FromLinear(double value)
    {
        value = Math.Clamp(value, 0, 1);
        return value <= 0.0031308 ? value * 12.92 : 1.055 * Math.Pow(value, 1 / 2.4) - 0.055;
    }

    private void TransformRgb(double scale, double offset)
    {
        this.red = FromLinear(ToLinear(this.red) * scale + offset);
        this.green = FromLinear(ToLinear(this.green) * scale + offset);
        this.blue = FromLinear(ToLinear(this.blue) * scale + offset);
    }

    private void TransformLuminance(double scale, double offset)
    {
        var min = Math.Min(this.red, Math.Min(this.green, this.blue));
        var max = Math.Max(this.red, Math.Max(this.green, this.blue));
        var luminance = (min + max) / 2;
        var chroma = max - min;
        var transformed = Math.Clamp(luminance * scale + offset, 0, 1);
        if (chroma == 0)
        {
            this.red = transformed;
            this.green = transformed;
            this.blue = transformed;
            return;
        }

        var saturation = chroma / (1 - Math.Abs(2 * luminance - 1));
        var newChroma = (1 - Math.Abs(2 * transformed - 1)) * saturation;
        var newMin = transformed - newChroma / 2;
        this.red = newMin + (this.red - min) / chroma * newChroma;
        this.green = newMin + (this.green - min) / chroma * newChroma;
        this.blue = newMin + (this.blue - min) / chroma * newChroma;
    }
}
