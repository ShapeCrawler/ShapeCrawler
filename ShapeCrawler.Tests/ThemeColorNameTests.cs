using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class ThemeColorNameTests : SCTest
{
    [TestCaseSource(nameof(ThemeColors))]
    public void ThemeColorName_returns_scheme_name(string schemeValue, string expectedName, bool tableCell)
    {
        // Arrange
        using var pres = PresentationWithThemeFill(schemeValue, tableCell);
        var fill = FillOf(pres, tableCell);
        var originalXml = pres.Slides[0].Shapes.Last().SdkOpenXmlElement.OuterXml;

        // Act
        var name = fill.ThemeColorName;

        // Assert
        name.Should().Be(expectedName);
        pres.Slides[0].Shapes.Last().SdkOpenXmlElement.OuterXml.Should().Be(originalXml,
            "reading the name must not change the fill");
        using var reopened = SaveAndOpenPresentation(pres);
        FillOf(reopened, tableCell).ThemeColorName.Should().Be(expectedName);
        ValidatePresentation(reopened);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void ThemeColorName_returns_null_after_replacing_theme_fill(bool tableCell)
    {
        // Arrange
        using var pres = PresentationWithThemeFill("accent2", tableCell);
        var fill = FillOf(pres, tableCell);
        fill.ThemeColorName.Should().Be("Accent 2");

        // Act-Assert
        using var image = TestAsset("09 png image.png");
        fill.SetPicture(image);
        fill.ThemeColorName.Should().BeNull();
        ValidatePresentation(pres);

        fill.SetColor("FF0000");
        fill.ThemeColorName.Should().BeNull();

        fill.SetNoFill();
        fill.ThemeColorName.Should().BeNull();
        ValidatePresentation(pres);
    }

    [Test]
    [SlideShape("009_table.pptx", 2, 6, null)]
    [SlideShape("009_table.pptx", 2, 2, null)]
    [SlideShape("009_table.pptx", 3, 4, null)]
    [SlideShape("009_table.pptx", 2, "AutoShape 1")]
    [SlideShape("009_table.pptx", 2, "AutoShape 2")]
    public void ThemeColorName_returns_null_for_fills_without_a_solid_theme_color(IShape shape)
    {
        // Act-Assert
        shape.Fill.ThemeColorName.Should().BeNull();
    }

    [Test]
    public void ThemeColorName_preserves_name_and_luminance_of_existing_fill()
    {
        // Arrange
        using var pres = new Presentation(TestAsset("009_table.pptx"));
        var fill = pres.Slides[3].Shapes.Shape("SolidSchemeLumDark").Fill;
        var originalColor = fill.Color;

        // Act-Assert
        fill.ThemeColorName.Should().Be("Accent 1");
        fill.LuminanceModulation.Should().Be(75);
        fill.Color.Should().Be(originalColor);
    }

    private static IEnumerable<TestCaseData> ThemeColors()
    {
        var colors = new (string Value, string Name)[]
        {
            ("bg1", "Background 1"),
            ("tx1", "Text 1"),
            ("bg2", "Background 2"),
            ("tx2", "Text 2"),
            ("dk1", "Dark 1"),
            ("lt1", "Light 1"),
            ("dk2", "Dark 2"),
            ("lt2", "Light 2"),
            ("accent1", "Accent 1"),
            ("accent2", "Accent 2"),
            ("accent3", "Accent 3"),
            ("accent4", "Accent 4"),
            ("accent5", "Accent 5"),
            ("accent6", "Accent 6"),
            ("hlink", "Hyperlink"),
            ("folHlink", "Followed Hyperlink"),
            ("phClr", "Placeholder")
        };

        foreach (var color in colors)
        {
            yield return new TestCaseData(color.Value, color.Name, false);
            yield return new TestCaseData(color.Value, color.Name, true);
        }
    }

    private static IShapeFill FillOf(Presentation pres, bool tableCell)
    {
        var shape = pres.Slides[0].Shapes.Last();
        return tableCell ? shape.Table![0, 0].Fill : shape.Fill;
    }

    private static Presentation PresentationWithThemeFill(string schemeValue, bool tableCell)
    {
        using var pres = new Presentation(p => p.Slide());
        var shapes = pres.Slides[0].Shapes;
        if (tableCell)
        {
            shapes.AddTable(0, 0, 1, 1);
        }
        else
        {
            shapes.AddShape(0, 0, 100, 100);
        }

        FillOf(pres, tableCell).SetColor("123456");
        using var stream = new MemoryStream();
        pres.Save(stream);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide!;
            OpenXmlCompositeElement properties = tableCell
                ? slide.Descendants<A.TableCellProperties>().First()
                : slide.Descendants<P.ShapeProperties>().First();
            var solidFill = properties.GetFirstChild<A.SolidFill>()!;
            solidFill.RemoveAllChildren();
            solidFill.Append(new A.SchemeColor(new A.LuminanceModulation { Val = 75000 })
            {
                Val = new EnumValue<A.SchemeColorValues> { InnerText = schemeValue }
            });
        }

        stream.Position = 0;
        return new Presentation(stream);
    }
}
