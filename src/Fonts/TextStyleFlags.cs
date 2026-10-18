using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Fonts;

internal sealed class TextStyleFlags(A.Text text)
{
    internal bool Bold => this.Resolve(properties => properties.Bold);

    internal bool Italic => this.Resolve(properties => properties.Italic);

    private static P.PlaceholderShape? Placeholder(P.Shape? shape) => shape?.NonVisualShapeProperties?
        .ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>();

    private static A.DefaultRunProperties? LevelProperties(OpenXmlElement? style, int level) => style?.Elements()
        .FirstOrDefault(element => element.LocalName == "lvl" + level.ToString(CultureInfo.InvariantCulture) + "pPr")?
        .GetFirstChild<A.DefaultRunProperties>();

    private static P.Shape? MasterPlaceholder(SlideMasterPart? master, P.PlaceholderValues type) => master?
        .SlideMaster?.CommonSlideData?.ShapeTree?.Elements<P.Shape>()
        .FirstOrDefault(shape => Placeholder(shape)?.Type?.Value == type);

    private bool Resolve(Func<A.TextCharacterPropertiesType, BooleanValue?> selector)
    {
        foreach (var properties in this.PropertyChain())
        {
            if (properties is not null && selector(properties)?.Value is { } value)
            {
                return value;
            }
        }

        return false;
    }

    private IEnumerable<A.TextCharacterPropertiesType?> PropertyChain()
    {
        yield return text.Parent?.GetFirstChild<A.RunProperties>();
        var paragraph = text.Ancestors<A.Paragraph>().First();
        yield return paragraph.ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>();
        var level = (paragraph.ParagraphProperties?.Level?.Value ?? 0) + 1;
        var body = paragraph.Parent;
        yield return LevelProperties(body?.GetFirstChild<A.ListStyle>(), level);

        var part = text.Ancestors<OpenXmlPartRootElement>().FirstOrDefault()?.OpenXmlPart;
        var shape = text.Ancestors<P.Shape>().FirstOrDefault();
        var placeholder = Placeholder(shape);
        if (placeholder is not null)
        {
            var type = placeholder.Type?.Value ?? P.PlaceholderValues.Object;
            SlideMasterPart? master = null;
            if (part is SlidePart slide)
            {
                var layout = slide.SlideLayoutPart;
                var layoutShape = layout?.SlideLayout?.CommonSlideData?.ShapeTree?.Elements<P.Shape>()
                    .FirstOrDefault(candidate => Placeholder(candidate) is { } reference &&
                        (reference.Index?.Value ?? 0) == (placeholder.Index?.Value ?? 0));
                yield return LevelProperties(layoutShape?.TextBody?.ListStyle, level);
                type = Placeholder(layoutShape)?.Type?.Value ?? type;
                master = layout?.SlideMasterPart;
            }
            else if (part is SlideLayoutPart layout)
            {
                master = layout.SlideMasterPart;
            }

            var masterType = type == P.PlaceholderValues.CenteredTitle ? P.PlaceholderValues.Title :
                type == P.PlaceholderValues.Object ? P.PlaceholderValues.Body : type;
            yield return LevelProperties(MasterPlaceholder(master, masterType)?.TextBody?.ListStyle, level);
            var styles = (master ?? part as SlideMasterPart)?.SlideMaster?.TextStyles;
            OpenXmlElement? style = masterType == P.PlaceholderValues.Title ? styles?.TitleStyle :
                masterType == P.PlaceholderValues.Body ? styles?.BodyStyle : styles?.OtherStyle;
            yield return LevelProperties(style, level);
        }

        var presentation = (part?.OpenXmlPackage as PresentationDocument)?.PresentationPart?.Presentation;
        yield return LevelProperties(presentation?.DefaultTextStyle, level);
    }
}
