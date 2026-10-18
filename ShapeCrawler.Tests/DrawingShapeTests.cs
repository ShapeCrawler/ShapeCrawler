using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class DrawingShapeTests : SCTest
{
    // Interior/exterior samples captured from Aspose.Slides 24.12, away from antialiased edges.
    [TestCase(Geometry.Triangle, "wide", "WWWWRWWWWWWWRRRWWWWW?RRR?WWW?RRRRR?W?RRRRRRR?")]
    [TestCase(Geometry.Triangle, "square", "WWWW?WWWWWWW?R?WWWWW?RRR?WWW?RRRRR?W?RRRRRRR?")]
    [TestCase(Geometry.Triangle, "tall", "WWWW?WWWWWWW?R?WWWWW?RRR?WWW?RRRRR?W?RRRRRRR?")]
    [TestCase(Geometry.Triangle, "adjusted", "WWW?WWWWWWWRRRWWWWW?RRRRWWWWRRRRRR?W?RRRRRRR?")]
    [TestCase(Geometry.Triangle, "flipped", "?RRRRRRR?W?RRRRR?WWW?RRR?WWWWWRRRWWWWWWWRWWWW")]
    [TestCase(Geometry.Triangle, "rotated", "WWWRRWWWWWWWRR?WWWWWWRRRWWWWWWRR?WWWWWWRRWWWW")]
    [TestCase(Geometry.RightTriangle, "wide", "?WWWWWWWWRR?WWWWWWRRRR?WWWWRRRRRR?WWRRRRRRRR?")]
    [TestCase(Geometry.RightTriangle, "square", "??WWWWWWWRR?WWWWWWRRRR?WWWWRRRRRR?WWRRRRRRR??")]
    [TestCase(Geometry.RightTriangle, "tall", "??WWWWWWWRR?WWWWWWRRRR?WWWWRRRRRR?WWRRRRRRR??")]
    [TestCase(Geometry.RightTriangle, "adjusted", "?WWWWWWWWRR?WWWWWWRRRR?WWWWRRRRRR?WWRRRRRRRR?")]
    [TestCase(Geometry.RightTriangle, "flipped", "?RRRRRRRRWW?RRRRRRWWWW?RRRRWWWWWW?RRWWWWWWWW?")]
    [TestCase(Geometry.RightTriangle, "rotated", "WWWRRWWWWWWWRRWWWWWWWR?WWWWWWWRWWWWWWWWRWWWWW")]
    [TestCase(Geometry.Diamond, "wide", "WWW?R?WWWW?RRRRR?WRRRRRRRRRW?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.Diamond, "square", "WWW?R?WWWW?RRRRR?W?RRRRRRR?W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.Diamond, "tall", "WWW?R?WWWW?RRRRR?W?RRRRRRR?W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.Diamond, "adjusted", "WWW?R?WWWW?RRRRR?WRRRRRRRRRW?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.Diamond, "flipped", "WWW?R?WWWW?RRRRR?WRRRRRRRRRW?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.Diamond, "rotated", "WWW?R?WWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWW?R?WWW")]
    [TestCase(Geometry.Parallelogram, "wide", "WRRRRRRRR?RRRRRRRRRRRRRRRR?RRRRRRRR?RRRRRRRRW")]
    [TestCase(Geometry.Parallelogram, "square", "WWRRRRRR?W?RRRRRR?WRRRRRR?W?RRRRRR?W?RRRRRRWW")]
    [TestCase(Geometry.Parallelogram, "tall", "WWRRRRRR?W?RRRRRR?WRRRRRRRW?RRRRRR?W?RRRRRRWW")]
    [TestCase(Geometry.Parallelogram, "adjusted", "WRRRRRRRRWRRRRRRRR?RRRRRRR?RRRRRRRRWRRRRRRRRW")]
    [TestCase(Geometry.Parallelogram, "flipped", "WRRRRRRRR?RRRRRRRRRRRRRRRR?RRRRRRRR?RRRRRRRRW")]
    [TestCase(Geometry.Parallelogram, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.Trapezoid, "wide", "WRRRRRRRW?RRRRRRR?RRRRRRRR?RRRRRRRRRRRRRRRRRR")]
    [TestCase(Geometry.Trapezoid, "square", "WWRRRRRWWW?RRRRR?WWRRRRRR?W?RRRRRRR??RRRRRRR?")]
    [TestCase(Geometry.Trapezoid, "tall", "WWRRRRRWWW?RRRRR?WWRRRRRRRW?RRRRRRR??RRRRRRR?")]
    [TestCase(Geometry.Trapezoid, "adjusted", "WRRRRRRRWWRRRRRRRW?RRRRRRR?RRRRRRRRRRRRRRRRRR")]
    [TestCase(Geometry.Trapezoid, "flipped", "RRRRRRRRRRRRRRRRRRRRRRRRRR??RRRRRRR?WRRRRRRRW")]
    [TestCase(Geometry.Trapezoid, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.Pentagon, "wide", "WWW?R?WWW??RRRRR???RRRRRRR?WRRRRRRRWW?RRRRR?W")]
    [TestCase(Geometry.Pentagon, "square", "WW??R??WW??RRRRR???RRRRRRR??RRRRRRR?W?RRRRR?W")]
    [TestCase(Geometry.Pentagon, "tall", "WWW?R?WWW??RRRRR???RRRRRRR??RRRRRRR?W?RRRRR?W")]
    [TestCase(Geometry.Pentagon, "adjusted", "WWW?R?WWW??RRRRR???RRRRRRR?WRRRRRRRWW?RRRRR?W")]
    [TestCase(Geometry.Pentagon, "flipped", "W?RRRRR?WWRRRRRRRW?RRRRRRR???RRRRR??WWW?R?WWW")]
    [TestCase(Geometry.Pentagon, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.Hexagon, "wide", "?RRRRRRR?RRRRRRRRRRRRRRRRRRRRRRRRRRR?RRRRRRR?")]
    [TestCase(Geometry.Hexagon, "square", "W?RRRRR?W?RRRRRRR??RRRRRRR??RRRRRRR?W?RRRRR?W")]
    [TestCase(Geometry.Hexagon, "tall", "W?RRRRR?W?RRRRRRR??RRRRRRR??RRRRRRR?W?RRRRR?W")]
    [TestCase(Geometry.Hexagon, "adjusted", "WRRRRRRRW?RRRRRRR?RRRRRRRRR?RRRRRRR?WRRRRRRRW")]
    [TestCase(Geometry.Hexagon, "flipped", "?RRRRRRR?RRRRRRRRRRRRRRRRRRRRRRRRRRR?RRRRRRR?")]
    [TestCase(Geometry.Hexagon, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.Octagon, "wide", "?RRRRRRR?RRRRRRRRRRRRRRRRRRRRRRRRRRR?RRRRRRR?")]
    [TestCase(Geometry.Octagon, "square", "W?RRRRR?W?RRRRRRR?RRRRRRRRR?RRRRRRR?W?RRRRR?W")]
    [TestCase(Geometry.Octagon, "tall", "W?RRRRR?WRRRRRRRRRRRRRRRRRRRRRRRRRRRW?RRRRR?W")]
    [TestCase(Geometry.Octagon, "adjusted", "WRRRRRRRWRRRRRRRRRRRRRRRRRRRRRRRRRRRWRRRRRRRW")]
    [TestCase(Geometry.Octagon, "flipped", "?RRRRRRR?RRRRRRRRRRRRRRRRRRRRRRRRRRR?RRRRRRR?")]
    [TestCase(Geometry.Octagon, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.Plus, "wide", "WRRRRRRRWRRRRRRRRRRRRRRRRRRRRRRRRRRRWRRRRRRRW")]
    [TestCase(Geometry.Plus, "square", "WW?RRR?WWRRRRRRRRRRRRRRRRRRRRRRRRRRRWW?RRR?WW")]
    [TestCase(Geometry.Plus, "tall", "WW?RRR?WWRRRRRRRRRRRRRRRRRRRRRRRRRRRWW?RRR?WW")]
    [TestCase(Geometry.Plus, "adjusted", "WRRRRRRRWWRRRRRRRWRRRRRRRRRWRRRRRRRWWRRRRRRRW")]
    [TestCase(Geometry.Plus, "flipped", "WRRRRRRRWRRRRRRRRRRRRRRRRRRRRRRRRRRRWRRRRRRRW")]
    [TestCase(Geometry.Plus, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.Chevron, "wide", "?RRRRRRRWWRRRRRRR?W?RRRRRRRWRRRRRRR??RRRRRRRW")]
    [TestCase(Geometry.Chevron, "square", "??RRR?WWWWW?RRRR?WWWWW?RRR?WW?RRRR?W??RRR?WWW")]
    [TestCase(Geometry.Chevron, "tall", "??RRR?WWWWW?RRRR?WWWWW?RRR?WW?RRRR?W??RRR?WWW")]
    [TestCase(Geometry.Chevron, "adjusted", "RRRRRRRRW?RRRRRRR?WRRRRRRRR?RRRRRRR?RRRRRRRRW")]
    [TestCase(Geometry.Chevron, "flipped", "WRRRRRRR??RRRRRRRWRRRRRRR?W?RRRRRRRWWRRRRRRR?")]
    [TestCase(Geometry.Chevron, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.HomePlate, "wide", "RRRRRRRRWRRRRRRRR?RRRRRRRRRRRRRRRRR?RRRRRRRRW")]
    [TestCase(Geometry.HomePlate, "square", "RRRRR?WWWRRRRRRR?WRRRRRRRR?RRRRRRR?WRRRRR?WWW")]
    [TestCase(Geometry.HomePlate, "tall", "RRRRR?WWWRRRRRRR?WRRRRRRRR?RRRRRRR?WRRRRR?WWW")]
    [TestCase(Geometry.HomePlate, "adjusted", "RRRRRRRRWRRRRRRRR?RRRRRRRRRRRRRRRRR?RRRRRRRRW")]
    [TestCase(Geometry.HomePlate, "flipped", "WRRRRRRRR?RRRRRRRRRRRRRRRRR?RRRRRRRRWRRRRRRRR")]
    [TestCase(Geometry.HomePlate, "rotated", "WWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.RightArrow, "wide", "WWWWWWW?WRRRRRRRR?RRRRRRRRRRRRRRRRR?WWWWWWW?W")]
    [TestCase(Geometry.RightArrow, "square", "WWWW??WWWRRRRRRR?WRRRRRRRR?RRRRRRR?WWWWW??WWW")]
    [TestCase(Geometry.RightArrow, "tall", "WWWW??WWWRRRRRRR?WRRRRRRRR?RRRRRRR?WWWWW??WWW")]
    [TestCase(Geometry.RightArrow, "adjusted", "WWWWWWWWWWWWWWWWW?RRRRRRRRRWWWWWWWW?WWWWWWWWW")]
    [TestCase(Geometry.RightArrow, "flipped", "W?WWWWWWW?RRRRRRRRRRRRRRRRR?RRRRRRRRW?WWWWWWW")]
    [TestCase(Geometry.RightArrow, "rotated", "WWWWRWWWWWWWWRWWWWWWWWRWWWWWWWWRWWWWWWWWRWWWW")]
    [TestCase(Geometry.LeftArrow, "wide", "W?WWWWWWW?RRRRRRRRRRRRRRRRR?RRRRRRRRW?WWWWWWW")]
    [TestCase(Geometry.LeftArrow, "square", "WWW??WWWWW?RRRRRRR?RRRRRRRRW?RRRRRRRWWW??WWWW")]
    [TestCase(Geometry.LeftArrow, "tall", "WWW??WWWWW?RRRRRRR?RRRRRRRRW?RRRRRRRWWW??WWWW")]
    [TestCase(Geometry.LeftArrow, "adjusted", "WWWWWWWWW?WWWWWWWWRRRRRRRRR?WWWWWWWWWWWWWWWWW")]
    [TestCase(Geometry.LeftArrow, "flipped", "WWWWWWW?WRRRRRRRR?RRRRRRRRRRRRRRRRR?WWWWWWW?W")]
    [TestCase(Geometry.LeftArrow, "rotated", "WWWWRWWWWWWWWRWWWWWWWWRWWWWWWWWRWWWWWWWWRWWWW")]
    [TestCase(Geometry.UpArrow, "wide", "WWW?R?WWWW?RRRRR?W??RRRRR??WWRRRRRWWWWRRRRRWW")]
    [TestCase(Geometry.UpArrow, "square", "WWW?R?WWWW?RRRRR?W???RRR???WW?RRR?WWWW?RRR?WW")]
    [TestCase(Geometry.UpArrow, "tall", "WWW?R?WWW??RRRRR??WW?RRR?WWWW?RRR?WWWW?RRR?WW")]
    [TestCase(Geometry.UpArrow, "adjusted", "WW??R??WW???RRR???WWWRRRWWWWWWRRRWWWWWWRRRWWW")]
    [TestCase(Geometry.UpArrow, "flipped", "WWRRRRRWWWWRRRRRWW??RRRRR??W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.UpArrow, "rotated", "WWWRR?WWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWWRR?WWW")]
    [TestCase(Geometry.DownArrow, "wide", "WWRRRRRWWWWRRRRRWW??RRRRR??W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.DownArrow, "square", "WW?RRR?WWWW?RRR?WW???RRR???W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.DownArrow, "tall", "WW?RRR?WWWW?RRR?WWWW?RRR?WW??RRRRR??WWW?R?WWW")]
    [TestCase(Geometry.DownArrow, "adjusted", "WWWRRRWWWWWWRRRWWWWWWRRRWWW???RRR???WW??R??WW")]
    [TestCase(Geometry.DownArrow, "flipped", "WWW?R?WWWW?RRRRR?W??RRRRR??WWRRRRRWWWWRRRRRWW")]
    [TestCase(Geometry.DownArrow, "rotated", "WWW?RRWWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWW?RRWWW")]
    [TestCase(Geometry.LeftRightArrow, "wide", "W?WWWWW?W?RRRRRRR?RRRRRRRRR?RRRRRRR?W?WWWWW?W")]
    [TestCase(Geometry.LeftRightArrow, "square", "WWW?R?WWWW?RRRRR?W?RRRRRRR?W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.LeftRightArrow, "tall", "WWW?R?WWWW?RRRRR?W?RRRRRRR?W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.LeftRightArrow, "adjusted", "WWWWWWWWW?WWWWWWW?RRRRRRRRR?WWWWWWW?WWWWWWWWW")]
    [TestCase(Geometry.LeftRightArrow, "flipped", "W?WWWWW?W?RRRRRRR?RRRRRRRRR?RRRRRRR?W?WWWWW?W")]
    [TestCase(Geometry.LeftRightArrow, "rotated", "WWWWRWWWWWWWWRWWWWWWWWRWWWWWWWWRWWWWWWWWRWWWW")]
    [TestCase(Geometry.UpDownArrow, "wide", "WWW?R?WWWW?RRRRR?WRRRRRRRRRW?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.UpDownArrow, "square", "WWW?R?WWWW?RRRRR?W?RRRRRRR?W?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.UpDownArrow, "tall", "WWW?R?WWW??RRRRR??WW?RRR?WW??RRRRR??WWW?R?WWW")]
    [TestCase(Geometry.UpDownArrow, "adjusted", "WW??R??WW???RRR???WWWRRRWWW???RRR???WW??R??WW")]
    [TestCase(Geometry.UpDownArrow, "flipped", "WWW?R?WWWW?RRRRR?WRRRRRRRRRW?RRRRR?WWWW?R?WWW")]
    [TestCase(Geometry.UpDownArrow, "rotated", "WWW?R?WWWWWWRRRWWWWWWRRRWWWWWWRRRWWWWWW?R?WWW")]
    public void Render_common_presets_match_reference_samples(Geometry geometry, string variant, string expected)
    {
        // Arrange
        var width = variant is "square" or "tall" ? 60 : 180;
        var height = variant == "tall" ? 80 : 60;
        using var pres = PresetPresentation(geometry, width, height, shape =>
        {
            if (variant == "adjusted")
            {
                var adjustments = shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.AdjustValueList!;
                foreach (var (name, value) in new[] { ("adj", 35000), ("adj1", 25000), ("adj2", 30000) })
                {
                    adjustments.Append(new A.ShapeGuide { Name = name, Formula = $"val {value}" });
                }
            }

            if (variant == "flipped")
            {
                shape.ShapeProperties!.Transform2D!.HorizontalFlip = true;
                shape.ShapeProperties.Transform2D.VerticalFlip = true;
            }

            if (variant == "rotated")
            {
                shape.ShapeProperties!.Transform2D!.Rotation = 90 * 60000;
            }
        });
        using var output = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(output, 300, 300);

        // Assert
        using var bitmap = SKBitmap.Decode(output);
        for (var row = 0; row < 5; row++)
        {
            for (var column = 0; column < 9; column++)
            {
                var sample = expected[row * 9 + column];
                if (sample == '?')
                {
                    continue;
                }

                var x = 50 + (int)Math.Round(width * (column + 0.5) / 9);
                var y = 200 + (int)Math.Round(height * (row + 0.5) / 5);
                bitmap.GetPixel(x, y).Should().Be(sample == 'R' ? SKColors.Red : SKColors.White,
                    $"{geometry}/{variant} at ({x}, {y})");
            }
        }

        ValidatePresentation(pres);
    }

    [TestCase(Geometry.Triangle, "adj", -100000, 65, 210)]
    [TestCase(Geometry.Triangle, "adj", 200000, 215, 210)]
    [TestCase(Geometry.RightArrow, "adj1", 200000, 65, 205)]
    public void Render_clamps_preset_adjustments(Geometry geometry, string name, int value, int x, int y)
    {
        // Arrange
        using var pres = PresetPresentation(geometry, 180, 60, shape =>
            shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.AdjustValueList!.Append(
                new A.ShapeGuide { Name = name, Formula = $"val {value}" }));
        using var output = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(output, 300, 300);

        // Assert
        using var bitmap = SKBitmap.Decode(output);
        bitmap.GetPixel(x, y).Should().Be(SKColors.Red);
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_preserves_preset_xml_and_scales_requested_output()
    {
        // Arrange
        using var pres = PresetPresentation(Geometry.Chevron, 180, 60, shape =>
            shape.ShapeProperties!.GetFirstChild<A.PresetGeometry>()!.AdjustValueList!.Append(
                new A.ShapeGuide { Name = "adj", Formula = "val 35000" }));
        using var before = new MemoryStream();
        pres.Save(before);
        using var output = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(output, 600, 600);

        // Assert
        using var bitmap = SKBitmap.Decode(output);
        bitmap.GetPixel(300, 460).Should().Be(SKColors.Red);
        bitmap.GetPixel(110, 460).Should().Be(SKColors.White);
        using var after = new MemoryStream();
        pres.Save(after);
        using var original = PresentationDocument.Open(before, false);
        using var saved = PresentationDocument.Open(after, false);
        saved.PresentationPart!.SlideParts.Single().Slide.OuterXml.Should()
            .Be(original.PresentationPart!.SlideParts.Single().Slide.OuterXml);
        ValidatePresentation(pres);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Render_respects_no_fill_and_solid_outline(bool noFill)
    {
        // Arrange
        using var pres = PresetPresentation(Geometry.Diamond, 180, 60, shape =>
        {
            var properties = shape.ShapeProperties!;
            if (noFill)
            {
                properties.RemoveAllChildren<A.SolidFill>();
                properties.InsertBefore(new A.NoFill(), properties.GetFirstChild<A.Outline>());
            }

            properties.RemoveAllChildren<A.Outline>();
            properties.Append(new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "0000FF" }))
                { Width = 6 * 12700 });
        });
        using var output = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(output, 300, 300);

        // Assert
        using var bitmap = SKBitmap.Decode(output);
        bitmap.GetPixel(140, 230).Should().Be(noFill ? SKColors.White : SKColors.Red);
        bitmap.GetPixel(140, 201).Should().Be(SKColors.Blue);
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_respects_solid_fill_alpha()
    {
        // Arrange
        using var pres = PresetPresentation(Geometry.Hexagon, 180, 60, shape =>
            shape.ShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Append(
                new A.Alpha { Val = 50000 }));
        using var output = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(output, 300, 300);

        // Assert
        using var bitmap = SKBitmap.Decode(output);
        bitmap.GetPixel(140, 230).Should().Be(new SKColor(255, 128, 128));
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_unsupported_geometry_still_reports_an_error()
    {
        // Arrange
        using var pres = PresetPresentation(Geometry.Star32, 180, 60);
        using var output = new MemoryStream();

        // Act
        var render = () => pres.Slide(1).SaveImage(output, 300, 300);

        // Assert
        render.Should().Throw<SCException>().WithMessage("Unsupported shape geometry type.*");
    }

    private static Presentation PresetPresentation(Geometry geometry, int width, int height,
        Action<P.Shape>? configure = null)
    {
        using var source = new Presentation(p => p.Slide());
        source.SlideWidth = 300;
        source.SlideHeight = 300;
        source.Slide(1).Shapes.AddShape(50, 200, width, height);
        var shape = source.Slide(1).Shapes.Last();
        shape.GeometryType = geometry;
        shape.Fill!.SetColor("FF0000");
        using var stream = new MemoryStream();
        source.Save(stream);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var sdkShape = document.PresentationPart!.SlideParts.Single().Slide.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>().Single();
            sdkShape.ShapeStyle?.Remove();
            sdkShape.ShapeProperties!.RemoveAllChildren<A.Outline>();
            sdkShape.ShapeProperties.Append(new A.Outline(new A.NoFill()));
            configure?.Invoke(sdkShape);
        }

        stream.Position = 0;
        return new Presentation(stream);
    }
}
