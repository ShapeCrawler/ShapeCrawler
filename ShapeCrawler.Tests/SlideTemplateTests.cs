using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class SlideTemplateTests : SCTest
{
    [TestCase(null, null, "FF0000", "008000")]
    [TestCase(false, null, "FFFFFF", "FFFFFF")]
    [TestCase(null, false, "FFFFFF", "008000")]
    [TestCase(false, false, "FFFFFF", "FFFFFF")]
    [TestCase(true, false, "FFFFFF", "008000")]
    [TestCase(false, true, "FFFFFF", "FFFFFF")]
    public void Render_orders_master_layout_and_slide_and_respects_visibility(
        bool? slideFlag, bool? layoutFlag, string masterPixel, string layoutPixel)
    {
        // Arrange
        using var pres = TemplatePresentation(slide =>
        {
            slide.Slide.ShowMasterShapes = slideFlag;
            slide.SlideLayoutPart!.SlideLayout.ShowMasterShapes = layoutFlag;
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(10, 70).Should().Be(SKColor.Parse(masterPixel));
        bitmap.GetPixel(30, 70).Should().Be(SKColor.Parse(layoutPixel));
        bitmap.GetPixel(50, 70).Should().Be(SKColors.Blue);
        ValidatePresentation(pres);
    }

    [TestCase("master", "FFFFFF", "008000", "0000FF")]
    [TestCase("layout", "FF0000", "FF0000", "0000FF")]
    [TestCase("slide", "FF0000", "008000", "008000")]
    public void Render_skips_hidden_shapes_on_each_layer(string layer, string left, string middle, string right)
    {
        // Arrange
        using var pres = TemplatePresentation(slide =>
        {
            var shape = LayerData(slide, layer).ShapeTree!.Elements<P.Shape>().Single(s => ShapeName(s) == layer);
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Hidden = true;
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(10, 70).Should().Be(SKColor.Parse(left));
        bitmap.GetPixel(30, 70).Should().Be(SKColor.Parse(middle));
        bitmap.GetPixel(50, 70).Should().Be(SKColor.Parse(right));
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_skips_template_placeholders_but_keeps_slide_placeholders()
    {
        // Arrange
        using var pres = TemplatePresentation(slide =>
        {
            foreach (var layer in new[] { "master", "layout" })
            {
                var data = LayerData(slide, layer);
                var placeholder = (P.Shape)data.ShapeTree!.Elements<P.Shape>().Single(s => ShapeName(s) == layer).CloneNode(true);
                placeholder.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = 500;
                placeholder.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties!.AppendChild(
                    new P.PlaceholderShape { Type = P.PlaceholderValues.Title });
                placeholder.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
                placeholder.ShapeProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" }));
                data.ShapeTree.Append(placeholder);
            }

            var own = slide.Slide.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single(s => ShapeName(s) == "slide");
            own.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.AppendChild(
                new P.PlaceholderShape { Type = P.PlaceholderValues.Title });
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(10, 70).Should().Be(SKColors.Red);
        bitmap.GetPixel(30, 70).Should().Be(SKColors.Green);
        bitmap.GetPixel(50, 70).Should().Be(SKColors.Blue);
        ValidatePresentation(pres);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Render_renders_template_picture_from_its_own_part(bool masterPicture)
    {
        // Arrange
        using var source = new SKBitmap(2, 1);
        source.SetPixel(0, 0, SKColors.Red);
        source.SetPixel(1, 0, SKColors.Blue);
        using var image = SKImage.FromBitmap(source);
        using var png = image.Encode(SKEncodedImageFormat.Png, 100);
        using var pres = TemplatePresentation(slide =>
        {
            var layout = slide.SlideLayoutPart!;
            var picture = new P.Picture(
                new P.NonVisualPictureProperties(new P.NonVisualDrawingProperties { Id = 600, Name = "Logo" },
                    new P.NonVisualPictureDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
                new P.BlipFill(new A.Blip { Embed = "logoImage" }, new A.Stretch(new A.FillRectangle())),
                new P.ShapeProperties(new A.Transform2D(new A.Offset { X = 0, Y = 127000 },
                        new A.Extents { Cx = 1270000, Cy = 254000 }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));
            using var imageStream = new MemoryStream(png.ToArray());
            var part = masterPicture ? layout.SlideMasterPart!.AddImagePart(ImagePartType.Png, "logoImage")
                : layout.AddImagePart(ImagePartType.Png, "logoImage");
            part.FeedData(imageStream);
            LayerData(slide, masterPicture ? "master" : "layout").ShapeTree!.Append(picture);
            // A slide-owned relationship with the same ID must not be used for a template picture.
            using var other = new SKBitmap(2, 1);
            other.Erase(SKColors.Green);
            using var otherImage = SKImage.FromBitmap(other);
            using var otherPng = otherImage.Encode(SKEncodedImageFormat.Png, 100);
            using var otherStream = new MemoryStream(otherPng.ToArray());
            slide.AddImagePart(ImagePartType.Png, "logoImage").FeedData(otherStream);
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(10, 20).Should().Be(SKColors.Red);
        bitmap.GetPixel(90, 20).Should().Be(SKColors.Blue);
        ValidatePresentation(pres);
    }

    [TestCase("master")]
    [TestCase("layout")]
    public void Render_renders_nonplaceholder_template_text(string layer)
    {
        // Arrange
        using var pres = TemplatePresentation(slide =>
        {
            var data = LayerData(slide, layer);
            var shape = (P.Shape)data.ShapeTree!.Elements<P.Shape>().Single(s => ShapeName(s) == layer).CloneNode(true);
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = 700;
            shape.ShapeProperties!.RemoveAllChildren<A.SolidFill>();
            shape.ShapeProperties.Append(new A.NoFill());
            shape.ShapeStyle?.Remove();
            shape.ShapeProperties.Transform2D!.Offset = new A.Offset { X = 127000, Y = 127000 };
            shape.ShapeProperties.Transform2D.Extents = new A.Extents { Cx = 1270000, Cy = 508000 };
            shape.TextBody = new P.TextBody(new A.BodyProperties(), new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.RunProperties(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "000000" }),
                        new A.LatinFont { Typeface = "DejaVu Sans" }) { FontSize = 1800 }, new A.Text("ACME"))));
            data.ShapeTree.Append(shape);
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 240, 180);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        var darkPixels = 0;
        for (var y = 10; y < 100; y++)
        {
            for (var x = 10; x < 230; x++)
            {
                var pixel = bitmap.GetPixel(x, y);
                if (pixel.Red < 100 && pixel.Green < 100 && pixel.Blue < 100)
                {
                    darkPixels++;
                }
            }
        }

        darkPixels.Should().BeGreaterThan(20);
        ValidatePresentation(pres);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Render_filters_placeholders_and_hidden_members_in_template_groups(bool nested)
    {
        // Arrange
        using var pres = TemplatePresentation(slide =>
        {
            var master = LayerData(slide, "master").ShapeTree!;
            var red = master.Elements<P.Shape>().Single(s => ShapeName(s) == "master");
            red.Remove();
            var placeholder = (P.Shape)red.CloneNode(true);
            placeholder.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = 801;
            placeholder.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties!.AppendChild(new P.PlaceholderShape());
            placeholder.ShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val = "FFFF00";
            var hidden = (P.Shape)red.CloneNode(true);
            hidden.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = 802;
            hidden.NonVisualShapeProperties.NonVisualDrawingProperties.Hidden = true;
            hidden.ShapeProperties!.GetFirstChild<A.SolidFill>()!.RgbColorModelHex!.Val = "000000";
            var group = Group(800, red, placeholder, hidden);
            master.Append(nested ? Group(803, group) : group);
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream, 120, 90);

        // Assert
        using var bitmap = SKBitmap.Decode(stream);
        bitmap.GetPixel(10, 70).Should().Be(SKColors.Red);
        bitmap.GetPixel(30, 70).Should().Be(SKColors.Green);
        ValidatePresentation(pres);
    }

    [Test]
    public void Render_preserves_slide_layout_and_master_xml()
    {
        // Arrange
        using var pres = TemplatePresentation(_ => { });
        using var before = SaveAndOpenPresentationAsSdk(pres);
        var original = LayerXml(before);
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(stream);

        // Assert
        using var after = SaveAndOpenPresentationAsSdk(pres);
        LayerXml(after).Should().Equal(original);
        ValidatePresentation(pres);
    }

    private static P.GroupShape Group(uint id, params OpenXmlElement[] shapes)
    {
        var group = new P.GroupShape(new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties { Id = id, Name = "Template group" },
                new P.NonVisualGroupShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(new A.TransformGroup(new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 1524000, Cy = 1143000 }, new A.ChildOffset { X = 0, Y = 0 },
                new A.ChildExtents { Cx = 1524000, Cy = 1143000 })));
        group.Append(shapes);
        return group;
    }

    private static P.CommonSlideData LayerData(SlidePart slide, string layer) => layer switch
    {
        "master" => slide.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!,
        "layout" => slide.SlideLayoutPart!.SlideLayout.CommonSlideData!,
        _ => slide.Slide.CommonSlideData!
    };

    private static string ShapeName(P.Shape shape) => shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!;

    private static string[] LayerXml(PresentationDocument document)
    {
        var slide = document.PresentationPart!.SlideParts.Single();
        return [slide.Slide.OuterXml, slide.SlideLayoutPart!.SlideLayout.OuterXml,
            slide.SlideLayoutPart.SlideMasterPart!.SlideMaster.OuterXml];
    }

    private static Presentation TemplatePresentation(Action<SlidePart> configure)
    {
        using var source = new Presentation(p => p.Slide());
        source.SlideWidth = 120;
        source.SlideHeight = 90;
        var shapes = source.Slide(1).Shapes;
        foreach (var (name, x, color) in new[] { ("master", 0, "FF0000"), ("layout", 20, "008000"), ("slide", 40, "0000FF") })
        {
            shapes.AddShape(x, 60, 80, 20);
            shapes.Last().Name = name;
            shapes.Last().Fill.SetColor(color);
        }

        using var stream = new MemoryStream();
        source.Save(stream);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true))
        {
            var slide = document.PresentationPart!.SlideParts.Single();
            foreach (var (layer, id) in new[] { ("master", 101u), ("layout", 201u) })
            {
                var shape = slide.Slide.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single(s => ShapeName(s) == layer);
                shape.Remove();
                shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = id;
                LayerData(slide, layer).ShapeTree!.Append(shape);
            }

            configure(slide);
        }

        stream.Position = 0;
        return new Presentation(stream);
    }
}
