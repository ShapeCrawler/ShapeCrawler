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
            SetStyle(slide.Slide.CommonSlideData!, 1001, new A.SchemeColor { Val = A.SchemeColorValues.Background1 });
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
