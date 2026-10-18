using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class SlideBackgroundTests : SCTest
{
    [TestCase("shade", "2C5C8C")]
    [TestCase("tint", "C0CDE2")]
    [TestCase("lum-mod", "204060")]
    [TestCase("lum-off", "A0C0E0")]
    [TestCase("lum-mod-off", "6699CD")]
    [TestCase("lum-off-mod", "578FC8")]
    [TestCase("alpha-half", "9FBFDF")]
    [TestCase("alpha-zero", "FFFFFF")]
    [TestCase("alpha-full", "4080C0")]
    [TestCase("alpha-mod", "B3CCE6")]
    [TestCase("alpha-off", "9FBFDF")]
    [TestCase("tint-shade", "8C96A6")]
    [TestCase("shade-tint", "BEC4D0")]
    [TestCase("lum-white", "FFFFFF")]
    [TestCase("lum-black", "000000")]
    [TestCase("alpha-clamp-full", "4080C0")]
    [TestCase("alpha-clamp-zero", "FFFFFF")]
    [TestCase("shade-zero", "000000")]
    [TestCase("tint-zero", "FFFFFF")]
    public void Render_applies_background_color_transforms_in_order(string operation, string expectedColor)
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            var color = new A.RgbColorModelHex { Val = "4080C0" };
            color.Append(Transforms(operation));
            slide.Slide.CommonSlideData!.InsertAt(new P.Background(new P.BackgroundProperties(new A.SolidFill(color))), 0);
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(100, 80).Should().Be(SKColor.Parse(expectedColor));
        ValidatePresentation(pres);
    }

    [TestCase("000000", "lum-off", "404040")]
    [TestCase("FFFFFF", "lum-mod", "808080")]
    [TestCase("808080", "lum-off", "C0C0C0")]
    public void Render_preserves_neutral_colors_during_luminance_transforms(
        string originalColor, string operation, string expectedColor)
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            var color = new A.RgbColorModelHex { Val = originalColor };
            color.Append(Transforms(operation));
            slide.Slide.CommonSlideData!.InsertAt(new P.Background(new P.BackgroundProperties(new A.SolidFill(color))), 0);
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(100, 80).Should().Be(SKColor.Parse(expectedColor));
        ValidatePresentation(pres);
    }

    [TestCase("scheme")]
    [TestCase("system")]
    [TestCase("inherited")]
    public void Render_transforms_resolved_scheme_system_and_inherited_backgrounds(string source)
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            var master = slide.SlideLayoutPart!.SlideMasterPart!;
            var accent = master.ThemePart!.Theme.ThemeElements!.ColorScheme!.GetFirstChild<A.Accent1Color>()!;
            accent.RemoveAllChildren();
            accent.Append(new A.RgbColorModelHex { Val = "4080C0" });
            OpenXmlElement color = source switch
            {
                "scheme" => new A.SchemeColor { Val = A.SchemeColorValues.Accent1 },
                "system" => new A.SystemColor { Val = A.SystemColorValues.WindowText, LastColor = "4080C0" },
                _ => new A.RgbColorModelHex { Val = "4080C0" }
            };
            color.Append(Transforms("lum-mod-off"));
            var data = source == "inherited" ? master.SlideMaster.CommonSlideData! : slide.Slide.CommonSlideData!;
            data.Background?.Remove();
            data.InsertAt(new P.Background(new P.BackgroundProperties(new A.SolidFill(color))), 0);
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(100, 80).Should().Be(SKColor.Parse("6699CD"));
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_applies_reference_color_transforms_before_placeholder_transforms()
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            var styles = slide.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme.ThemeElements!.FormatScheme!.BackgroundFillStyleList!;
            styles.ChildElements[0].Remove();
            styles.InsertAt(new A.SolidFill(new A.SchemeColor(new A.Shade { Val = 50000 })
            {
                Val = A.SchemeColorValues.PhColor
            }), 0);
            SetStyle(slide.Slide.CommonSlideData!, 1001, new A.RgbColorModelHex(new A.Tint { Val = 50000 })
            {
                Val = "4080C0"
            });
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(100, 80).Should().Be(SKColor.Parse("8C96A6"));
        ValidatePresentation(pres);
    }

    [TestCase(0, "FFFFFF")]
    [TestCase(1, "FF0000")]
    [TestCase(2, "CC00CC")]
    [TestCase(3, "0000FF")]
    [TestCase(99, "FFFFFF")]
    [TestCase(1000, "FFFFFF")]
    [TestCase(1001, "CC00CC")]
    [TestCase(1002, "008000")]
    [TestCase(1003, "778899")]
    [TestCase(1004, "FFFFFF")]
    public void Render_selects_style_matrix_fill_and_substitutes_placeholder(int index, string expectedColor)
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
            SetStyle(slide.Slide.CommonSlideData!, index, new A.RgbColorModelHex { Val = "CC00CC" }));
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(100, 80).Should().Be(SKColor.Parse(expectedColor));
        ValidatePresentation(pres);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Render_uses_inherited_theme_background_style(bool masterBackground)
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            var layout = slide.SlideLayoutPart!;
            var data = masterBackground
                ? layout.SlideMasterPart!.SlideMaster.CommonSlideData!
                : layout.SlideLayout.CommonSlideData!;
            SetStyle(data, 1002, new A.RgbColorModelHex { Val = "CC00CC" });
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(100, 80).Should().Be(SKColors.Green);
        ValidatePresentation(pres);
    }

    [TestCase("slide", "layout", false, "445566")]
    [TestCase("layout", "layout", false, "445566")]
    [TestCase("master", "layout", false, "445566")]
    [TestCase("slide", "slide", false, "778899")]
    [TestCase("master", "slide", false, "778899")]
    [TestCase("slide", "master", false, "112233")]
    [TestCase("slide", "layout-master", false, "112233")]
    [TestCase("slide", "slide", true, "778899")]
    [TestCase("master", "slide", true, "778899")]
    [TestCase("slide", "master", true, "112233")]
    public void Render_resolves_effective_background_color_map(
        string backgroundLevel, string mapping, bool reference, string expectedColor)
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            SetColorMaps(slide, mapping);
            var layout = slide.SlideLayoutPart!;
            var data = backgroundLevel switch
            {
                "master" => layout.SlideMasterPart!.SlideMaster.CommonSlideData!,
                "layout" => layout.SlideLayout.CommonSlideData!,
                _ => slide.Slide.CommonSlideData!
            };
            if (reference)
            {
                SetStyle(data, 1001, new A.SchemeColor { Val = A.SchemeColorValues.Background1 });
            }
            else
            {
                data.Background?.Remove();
                data.InsertAt(new P.Background(new P.BackgroundProperties(new A.SolidFill(
                    new A.SchemeColor { Val = A.SchemeColorValues.Background1 }))), 0);
            }
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(100, 80).Should().Be(SKColor.Parse(expectedColor));
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_resolves_theme_style_picture_relationship_in_theme_part()
    {
        // Arrange
        using var bitmap = new SKBitmap(2, 2);
        bitmap.Erase(SKColors.Blue);
        using var image = SKImage.FromBitmap(bitmap);
        using var png = image.Encode(SKEncodedImageFormat.Png, 100);
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            var theme = slide.SlideLayoutPart!.SlideMasterPart!.ThemePart!;
            using var data = new MemoryStream(png.ToArray());
            theme.AddImagePart(ImagePartType.Png, "backgroundImage").FeedData(data);
            var styles = theme.Theme.ThemeElements!.FormatScheme!.BackgroundFillStyleList!;
            styles.ChildElements[1].Remove();
            styles.InsertAt(new A.BlipFill(new A.Blip { Embed = "backgroundImage" },
                new A.Stretch(new A.FillRectangle())), 1);
            SetStyle(slide.Slide.CommonSlideData!, 1002, new A.RgbColorModelHex { Val = "CC00CC" });
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 240, 45);

        // Assert
        using var actual = SKBitmap.Decode(stream);
        actual.GetPixel(2, 2).Should().Be(SKColors.Blue);
        actual.GetPixel(237, 42).Should().Be(SKColors.Blue);
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_preserves_background_reference_theme_and_color_maps()
    {
        // Arrange
        using var pres = ThemeBackgroundPresentation(slide =>
        {
            SetColorMaps(slide, "slide");
            SetStyle(slide.Slide.CommonSlideData!, 1001, new A.SchemeColor(new A.Tint { Val = 50000 })
            {
                Val = A.SchemeColorValues.Background1
            });
        });
        using var before = SaveAndOpenPresentationAsSdk(pres);
        var original = BackgroundXml(before);
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream);

        // Assert
        using var after = SaveAndOpenPresentationAsSdk(pres);
        BackgroundXml(after).Should().Equal(original);
        ValidatePresentation(pres);
    }

    private static OpenXmlElement[] Transforms(string operation) => operation switch
    {
        "shade" => [new A.Shade { Val = 50000 }],
        "tint" => [new A.Tint { Val = 50000 }],
        "lum-mod" => [new A.LuminanceModulation { Val = 50000 }],
        "lum-off" => [new A.LuminanceOffset { Val = 25000 }],
        "lum-mod-off" => [new A.LuminanceModulation { Val = 80000 }, new A.LuminanceOffset { Val = 20000 }],
        "lum-off-mod" => [new A.LuminanceOffset { Val = 20000 }, new A.LuminanceModulation { Val = 80000 }],
        "alpha-half" => [new A.Alpha { Val = 50000 }],
        "alpha-zero" => [new A.Alpha { Val = 0 }],
        "alpha-full" => [new A.Alpha { Val = 100000 }],
        "alpha-mod" => [new A.Alpha { Val = 80000 }, new A.AlphaModulation { Val = 50000 }],
        "alpha-off" => [new A.Alpha { Val = 25000 }, new A.AlphaOffset { Val = 25000 }],
        "tint-shade" => [new A.Tint { Val = 50000 }, new A.Shade { Val = 50000 }],
        "shade-tint" => [new A.Shade { Val = 50000 }, new A.Tint { Val = 50000 }],
        "lum-white" => [new A.LuminanceModulation { Val = 200000 }],
        "lum-black" => [new A.LuminanceOffset { Val = -100000 }],
        "alpha-clamp-full" => [new A.Alpha { Val = 80000 }, new A.AlphaModulation { Val = 200000 }],
        "alpha-clamp-zero" => [new A.Alpha { Val = 25000 }, new A.AlphaOffset { Val = -50000 }],
        "shade-zero" => [new A.Shade { Val = 0 }],
        "tint-zero" => [new A.Tint { Val = 0 }],
        _ => throw new ArgumentOutOfRangeException(nameof(operation))
    };

    private static string[] BackgroundXml(PresentationDocument document)
    {
        var slide = document.PresentationPart!.SlideParts.Single();
        var layout = slide.SlideLayoutPart!;
        var master = layout.SlideMasterPart!;
        return [slide.Slide.OuterXml, layout.SlideLayout.OuterXml, master.SlideMaster.OuterXml, master.ThemePart!.Theme.OuterXml];
    }

    private static Presentation ThemeBackgroundPresentation(Action<SlidePart> configure)
    {
        using var source = new Presentation(p => p.Slide());
        using var stream = new MemoryStream();
        source.Save(stream);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true))
        {
            var slide = document.PresentationPart!.SlideParts.Single();
            var layout = slide.SlideLayoutPart!;
            var master = layout.SlideMasterPart!;
            slide.Slide.CommonSlideData!.Background?.Remove();
            layout.SlideLayout.CommonSlideData!.Background?.Remove();
            master.SlideMaster.CommonSlideData!.Background?.Remove();
            master.SlideMaster.CommonSlideData.InsertAt(new P.Background(new P.BackgroundProperties(SolidFill("0000FF"))), 0);
            var elements = master.ThemePart!.Theme.ThemeElements!;
            foreach (var (color, value) in new[] { ("accent1", "112233"), ("accent2", "445566"), ("accent3", "778899") })
            {
                var entry = elements.ColorScheme!.ChildElements.Single(element => element.LocalName == color);
                entry.RemoveAllChildren();
                entry.AppendChild(new A.RgbColorModelHex { Val = value });
            }

            elements.FormatScheme!.FillStyleList = new A.FillStyleList(
                SolidFill("FF0000"), new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }), SolidFill("0000FF"));
            elements.FormatScheme.BackgroundFillStyleList = new A.BackgroundFillStyleList(
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.PhColor }), SolidFill("008000"),
                new A.SolidFill(new A.SchemeColor { Val = A.SchemeColorValues.Accent3 }));
            configure(slide);
        }

        stream.Position = 0;
        return new Presentation(stream);
    }

    private static A.SolidFill SolidFill(string color) => new(new A.RgbColorModelHex { Val = color });

    private static void SetStyle(P.CommonSlideData data, int index, OpenXmlElement color)
    {
        data.Background?.Remove();
        data.InsertAt(new P.Background(new P.BackgroundStyleReference(color) { Index = (uint)index }), 0);
    }

    private static void SetColorMaps(SlidePart slide, string mapping)
    {
        var layout = slide.SlideLayoutPart!;
        var master = layout.SlideMasterPart!;
        master.SlideMaster.ColorMap!.Background1 = A.ColorSchemeIndexValues.Accent1;
        var layoutMap = new A.OverrideColorMapping();
        layoutMap.SetAttributes(master.SlideMaster.ColorMap.GetAttributes());
        layoutMap.Background1 = A.ColorSchemeIndexValues.Accent2;
        layout.SlideLayout.ColorMapOverride = new P.ColorMapOverride(layoutMap);
        slide.Slide.ColorMapOverride?.Remove();
        if (mapping == "master")
        {
            slide.Slide.ColorMapOverride = new P.ColorMapOverride(new A.MasterColorMapping());
        }
        else if (mapping == "layout-master")
        {
            layout.SlideLayout.ColorMapOverride = new P.ColorMapOverride(new A.MasterColorMapping());
        }
        else if (mapping == "slide")
        {
            var slideMap = new A.OverrideColorMapping();
            slideMap.SetAttributes(master.SlideMaster.ColorMap.GetAttributes());
            slideMap.Background1 = A.ColorSchemeIndexValues.Accent3;
            slide.Slide.ColorMapOverride = new P.ColorMapOverride(slideMap);
        }
    }
}
