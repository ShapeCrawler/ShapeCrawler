using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Drawing;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

internal sealed class SlideBackground(SlidePart slidePart)
{
    internal void Render(SKCanvas canvas)
    {
        var layoutPart = slidePart.SlideLayoutPart;
        var background = slidePart.Slide?.CommonSlideData?.Background
                         ?? layoutPart?.SlideLayout?.CommonSlideData?.Background
                         ?? layoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background;
        var reference = background?.BackgroundStyleReference;
        var fill = background?.BackgroundProperties?.FirstChild ?? this.StyleFill(reference);

        switch (fill)
        {
            case A.SolidFill solid:
                canvas.Clear(this.ResolveColor(solid.FirstChild, reference?.FirstChild) ?? SKColors.White);
                break;
            case A.BlipFill { Blip: { Embed: not null } blip }:
                using (var bitmap = SKBitmap.Decode(new SlidePictureImage(blip).AsByteArray()))
                {
                    if (bitmap is null)
                    {
                        canvas.Clear(SKColors.White);
                        break;
                    }

                    var bounds = new SKRect(0, 0, canvas.DeviceClipBounds.Width, canvas.DeviceClipBounds.Height);
                    canvas.DrawBitmap(bitmap, bounds, SKSamplingOptions.Default);
                }

                break;
            default:
                canvas.Clear(SKColors.White);
                break;
        }
    }

    private static SKColor? ParseColor(string? hex)
    {
        hex = hex?.TrimStart('#');
        return hex?.Length is 6 or 8 ? new Color(hex).AsSkColor() : null;
    }

    private OpenXmlElement? StyleFill(P.BackgroundStyleReference? reference)
    {
        var index = reference?.Index?.Value ?? 0;
        if (index is 0 or 1000)
        {
            return null;
        }

        var format = slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.FormatScheme;
        OpenXmlElement? styles = index < 1000 ? format?.FillStyleList : format?.BackgroundFillStyleList;
        var offset = index < 1000 ? index - 1 : index - 1001;
        return styles is not null && offset < styles.ChildElements.Count ? styles.ChildElements[(int)offset] : null;
    }

    private SKColor? ResolveColor(OpenXmlElement? color, OpenXmlElement? placeholder)
    {
        switch (color)
        {
            case A.RgbColorModelHex rgb:
                return ParseColor(rgb.Val?.Value);
            case A.SystemColor system:
                return ParseColor(system.LastColor?.Value);
            case A.SchemeColor scheme when scheme.Val?.Value == A.SchemeColorValues.PhColor:
                return placeholder is A.SchemeColor { Val.Value: var value } && value == A.SchemeColorValues.PhColor
                    ? null
                    : this.ResolveColor(placeholder, null);
            case A.SchemeColor scheme:
                var name = this.MapSchemeColor(scheme.Val?.InnerText);
                var colors = slidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
                var entry = colors?.ChildElements.FirstOrDefault(element => element.LocalName == name);
                return entry?.FirstChild switch
                {
                    A.RgbColorModelHex rgb => ParseColor(rgb.Val?.Value),
                    A.SystemColor system => ParseColor(system.LastColor?.Value),
                    _ => null
                };
            default:
                return null;
        }
    }

    private string? MapSchemeColor(string? name)
    {
        var layout = slidePart.SlideLayoutPart;
        var masterMap = layout?.SlideMasterPart?.SlideMaster?.ColorMap;
        var slideOverride = slidePart.Slide?.ColorMapOverride;
        OpenXmlElement? mapping;
        if (slideOverride?.MasterColorMapping is not null)
        {
            mapping = masterMap;
        }
        else
        {
            mapping = slideOverride?.OverrideColorMapping
                      ?? (OpenXmlElement?)layout?.SlideLayout?.ColorMapOverride?.OverrideColorMapping
                      ?? masterMap;
        }

        var mapped = mapping?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name).Value;
        if (!string.IsNullOrEmpty(mapped))
        {
            return mapped;
        }

        return name switch
        {
            "bg1" => "lt1",
            "tx1" => "dk1",
            "bg2" => "lt2",
            "tx2" => "dk2",
            _ => name
        };
    }
}
