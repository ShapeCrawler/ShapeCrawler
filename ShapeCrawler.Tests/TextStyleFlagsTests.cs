using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using SkiaSharp;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class TextStyleFlagsTests : SCTest
{
    [TestCase("plain", false, false)]
    [TestCase("paragraph", true, true)]
    [TestCase("list", true, true)]
    [TestCase("list-default", false, false)]
    [TestCase("run-false", false, false)]
    [TestCase("paragraph-false", false, false)]
    [TestCase("layout", true, true)]
    [TestCase("master-shape", true, true)]
    [TestCase("master-style", true, true)]
    [TestCase("presentation", true, true)]
    [TestCase("mixed", false, false)]
    [TestCase("level2", true, true)]
    [TestCase("skip-unset", true, true)]
    [TestCase("no-run", true, true)]
    [TestCase("layout-idx", true, true)]
    [TestCase("layout-object", true, true)]
    [TestCase("title", true, true)]
    [TestCase("end-paragraph", true, true)]
    [TestCase("layout-para", false, false)]
    [TestCase("master-para", false, false)]
    [TestCase("master-object", true, true)]
    [TestCase("master-title", true, true)]
    [TestCase("other-style", false, false)]
    [TestCase("level-fallback", false, false)]
    public void Flags_follow_reference_precedence(string mode, bool bold, bool italic)
    {
        // Arrange
        using var pres = StyledPresentation(mode);
        var font = Font(pres);
        using var before = new MemoryStream();
        pres.Save(before);

        // Act
        var actual = (font.IsBold, font.IsItalic);

        // Assert
        actual.Should().Be((bold, italic));
        using var after = new MemoryStream();
        pres.Save(after);
        AssertXmlUnchanged(before, after);
        ValidatePresentation(pres);
    }

    [TestCase("paragraph", true, true)]
    [TestCase("list", true, true)]
    [TestCase("layout", true, true)]
    [TestCase("master-shape", true, true)]
    [TestCase("master-style", true, true)]
    [TestCase("presentation", true, true)]
    [TestCase("mixed", false, false)]
    public void Render_inherited_style_matches_explicit_style(string mode, bool bold, bool italic)
    {
        // Arrange
        using var pres = StyledPresentation(mode);
        using var before = new MemoryStream();
        pres.Save(before);
        using var inherited = new MemoryStream();

        // Act
        pres.Slide(1).SaveImage(inherited, 300, 300);

        // Assert
        using var after = new MemoryStream();
        pres.Save(after);
        AssertXmlUnchanged(before, after);
        var font = Font(pres);
        font.IsBold = bold;
        font.IsItalic = italic;
        using var explicitStyle = new MemoryStream();
        pres.Slide(1).SaveImage(explicitStyle, 300, 300);
        using var expected = SKBitmap.Decode(explicitStyle);
        using var actual = SKBitmap.Decode(inherited);
        actual.Pixels.Should().Equal(expected.Pixels);
        actual.Pixels.Count(pixel => pixel.Red < 100 && pixel.Green < 100 && pixel.Blue < 100)
            .Should().BeGreaterThan(30);
        ValidatePresentation(pres);
    }

    [TestCase(true, true)]
    [TestCase(false, false)]
    [TestCase(true, false)]
    [TestCase(false, true)]
    public void Setters_override_inherited_flags_on_only_the_selected_run(bool bold, bool italic)
    {
        // Arrange
        using var pres = StyledPresentation("no-run", shape =>
        {
            var paragraph = shape.TextBody!.Elements<A.Paragraph>().Single();
            paragraph.RemoveAllChildren<A.EndParagraphRunProperties>();
            paragraph.Append(new A.Run(new A.Text("Second")),
                new A.EndParagraphRunProperties { Bold = true, Italic = true });
        });
        var first = Font(pres);
        var second = pres.Slide(1).Shapes.Single().TextBox!.Paragraphs[0].Portions[1].Font!;

        // Act
        first.IsBold = bold;
        first.IsItalic = italic;

        // Assert
        (first.IsBold, first.IsItalic).Should().Be((bold, italic));
        (second.IsBold, second.IsItalic).Should().Be((true, true));
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var paragraph = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Paragraph>().Single();
            var run = paragraph.Elements<A.Run>().First();
            run.RunProperties!.Bold!.Value.Should().Be(bold);
            run.RunProperties.Italic!.Value.Should().Be(italic);
            paragraph.GetFirstChild<A.EndParagraphRunProperties>()!.Bold!.Value.Should().BeTrue();
            paragraph.GetFirstChild<A.EndParagraphRunProperties>()!.Italic!.Value.Should().BeTrue();
        }

        using var reopened = new Presentation(saved);
        (Font(reopened).IsBold, Font(reopened).IsItalic).Should().Be((bold, italic));
        ValidatePresentation(pres);
    }

    [TestCase(false, null, false, true)]
    [TestCase(null, false, true, false)]
    [TestCase(true, null, true, true)]
    [TestCase(null, true, true, true)]
    public void Flags_resolve_each_attribute_independently(bool? bold, bool? italic, bool expectedBold, bool expectedItalic)
    {
        // Arrange
        using var pres = StyledPresentation("list", shape => shape.TextBody!.Elements<A.Paragraph>().Single()
            .ParagraphProperties!.Append(new A.DefaultRunProperties { Bold = bold, Italic = italic }));
        var font = Font(pres);

        // Act
        var actual = (font.IsBold, font.IsItalic);

        // Assert
        actual.Should().Be((expectedBold, expectedItalic));
        ValidatePresentation(pres);
    }

    [TestCase(false, false)]
    [TestCase(true, true)]
    public void Flags_match_layout_placeholders_by_index_with_zero_as_default(bool defaultIndex, bool expected)
    {
        // Arrange
        using var pres = StyledPresentation("layout", shape =>
        {
            var slide = (SlidePart)shape.Ancestors<OpenXmlPartRootElement>().Single().OpenXmlPart!;
            var reference = slide.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
                .Single().NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>()!;
            reference.Index = defaultIndex ? null : 8;
            if (defaultIndex)
            {
                shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>()!
                    .Index = null;
            }
        });
        var font = Font(pres);

        // Act
        var actual = (font.IsBold, font.IsItalic);

        // Assert
        actual.Should().Be((expected, expected));
        ValidatePresentation(pres);
    }

    [TestCase(false)]
    [TestCase(true)]
    public void Flags_resolve_local_list_styles_for_template_text(bool masterText)
    {
        // Arrange
        using var pres = StyledPresentation("list", shape =>
        {
            var slide = (SlidePart)shape.Ancestors<OpenXmlPartRootElement>().Single().OpenXmlPart!;
            var tree = masterText ? slide.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!.ShapeTree!
                : slide.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!;
            var copy = (P.Shape)shape.CloneNode(true);
            copy.NonVisualShapeProperties!.NonVisualDrawingProperties!.Id = 300;
            copy.NonVisualShapeProperties.NonVisualDrawingProperties.Name = "Template text";
            tree.Append(copy);
        });
        var shapes = masterText ? pres.Slide(1).LayoutSlide.MasterSlide.Shapes : pres.Slide(1).LayoutSlide.Shapes;
        var font = shapes.Single(shape => shape.Name == "Template text").TextBox!.Paragraphs[0].Portions[0].Font!;

        // Act
        var actual = (font.IsBold, font.IsItalic);

        // Assert
        actual.Should().Be((true, true));
        ValidatePresentation(pres);
    }

    private static ITextPortionFont Font(Presentation pres) =>
        pres.Slide(1).Shapes.Single().TextBox!.Paragraphs[0].Portions[0].Font!;

    private static void AssertXmlUnchanged(Stream before, Stream after)
    {
        using var original = PresentationDocument.Open(before, false);
        using var saved = PresentationDocument.Open(after, false);
        var oldSlide = original.PresentationPart!.SlideParts.Single();
        var newSlide = saved.PresentationPart!.SlideParts.Single();
        newSlide.Slide.OuterXml.Should().Be(oldSlide.Slide.OuterXml);
        newSlide.SlideLayoutPart!.SlideLayout.OuterXml.Should().Be(oldSlide.SlideLayoutPart!.SlideLayout.OuterXml);
        newSlide.SlideLayoutPart.SlideMasterPart!.SlideMaster.OuterXml.Should()
            .Be(oldSlide.SlideLayoutPart.SlideMasterPart!.SlideMaster.OuterXml);
        saved.PresentationPart.Presentation.OuterXml.Should().Be(original.PresentationPart.Presentation.OuterXml);
    }

    private static Presentation StyledPresentation(string mode, Action<P.Shape>? configure = null)
    {
        using var stream = new MemoryStream();
        using (var source = new Presentation(p => p.Slide()))
        {
            source.SlideWidth = 300;
            source.SlideHeight = 300;
            source.Slide(1).Shapes.AddShape(20, 200, 260, 80);
            var shape = source.Slide(1).Shapes.Last();
            shape.SetText("MMMM Test");
            shape.SetFontSize(24);
            shape.SetFontName("Arial");
            shape.Fill!.SetNoFill();
            shape.Outline!.SetNoOutline();
            source.Save(stream);
        }

        using (var sdk = PresentationDocument.Open(stream, true))
        {
            var part = sdk.PresentationPart!;
            var slide = part.SlideParts.Single();
            var shape = slide.Slide.CommonSlideData!.ShapeTree!.Elements<P.Shape>().Single();
            var body = shape.TextBody!;
            var paragraph = body.Elements<A.Paragraph>().Single();
            var run = paragraph.Elements<A.Run>().Single();
            run.RunProperties!.Bold = null;
            run.RunProperties.Italic = null;
            var properties = paragraph.ParagraphProperties ?? paragraph.PrependChild(new A.ParagraphProperties());
            properties.Level = mode == "level2" ? 1 : 0;
            var level = properties.Level.Value;
            var layout = slide.SlideLayoutPart!.SlideLayout;
            var master = slide.SlideLayoutPart.SlideMasterPart!.SlideMaster;
            master.CommonSlideData!.ShapeTree!.RemoveAllChildren<P.Shape>();
            layout.CommonSlideData!.ShapeTree!.RemoveAllChildren<P.Shape>();
            var type = mode is "title" or "master-title" ? P.PlaceholderValues.Title :
                mode is "layout-object" or "master-object" ? P.PlaceholderValues.Object : P.PlaceholderValues.Body;
            if (new[] { "layout", "master-shape", "master-style", "mixed", "level2", "skip-unset", "no-run",
                    "layout-idx", "layout-object", "title", "run-false", "layout-para", "master-para",
                    "master-object", "master-title", "level-fallback" }.Contains(mode))
            {
                shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.Append(
                    new P.PlaceholderShape { Type = type, Index = 7 });
                var layoutShape = TemplatePlaceholder(101, 7, mode == "layout-idx" ? P.PlaceholderValues.Object : type);
                var masterShape = TemplatePlaceholder(201, 1, type == P.PlaceholderValues.Object ? P.PlaceholderValues.Body : type);
                layout.CommonSlideData.ShapeTree.Append(layoutShape);
                master.CommonSlideData.ShapeTree.Append(masterShape);
                if (new[] { "layout", "mixed", "level2", "skip-unset", "no-run", "layout-idx", "layout-object",
                        "title", "run-false" }.Contains(mode))
                {
                    SetLevel(layoutShape.TextBody!.ListStyle!, level, Flags());
                }

                if (mode == "layout-para")
                {
                    layoutShape.TextBody!.Elements<A.Paragraph>().Single().PrependChild(new A.ParagraphProperties(Flags()));
                }

                if (mode == "master-para")
                {
                    masterShape.TextBody!.Elements<A.Paragraph>().Single().PrependChild(new A.ParagraphProperties(Flags()));
                }

                if (mode is "master-shape" or "master-object")
                {
                    SetLevel(masterShape.TextBody!.ListStyle!, level, Flags());
                }

                if (mode == "master-title")
                {
                    SetLevel(master.TextStyles!.TitleStyle!, level, Flags());
                }

                if (mode == "master-style")
                {
                    SetLevel(master.TextStyles!.BodyStyle!, level, Flags());
                }

                if (mode == "level-fallback")
                {
                    properties.Level = 1;
                    SetLevel(body.ListStyle!, 0, Flags());
                }

                if (mode == "skip-unset")
                {
                    SetLevel(body.ListStyle!, level, new A.DefaultRunProperties { FontSize = 2400 });
                }

                if (mode == "mixed")
                {
                    SetLevel(body.ListStyle!, level, Flags(false, null));
                    run.RunProperties.Italic = false;
                }

                if (mode == "no-run")
                {
                    run.RunProperties.Remove();
                }
            }

            if (mode is "paragraph" or "paragraph-false")
            {
                properties.Append(Flags(mode == "paragraph", mode == "paragraph"));
            }

            if (mode is "list" or "paragraph-false" or "run-false" or "end-paragraph")
            {
                SetLevel(body.ListStyle!, level, Flags());
            }

            if (mode == "list-default")
            {
                body.ListStyle!.RemoveAllChildren();
                body.ListStyle.Append(new A.DefaultParagraphProperties(Flags()));
            }

            if (mode == "run-false")
            {
                run.RunProperties.Bold = false;
                run.RunProperties.Italic = false;
            }

            if (mode == "other-style")
            {
                SetLevel(master.TextStyles!.OtherStyle!, level, Flags());
            }

            if (mode == "presentation")
            {
                SetLevel(part.Presentation.DefaultTextStyle!, level, Flags());
            }

            if (mode == "end-paragraph")
            {
                paragraph.RemoveAllChildren<A.EndParagraphRunProperties>();
                paragraph.Append(new A.EndParagraphRunProperties { Bold = false, Italic = false });
            }

            configure?.Invoke(shape);
        }

        stream.Position = 0;
        return new Presentation(stream);
    }

    private static A.DefaultRunProperties Flags(bool? bold = true, bool? italic = true) =>
        new() { Bold = bold, Italic = italic };

    private static void SetLevel(OpenXmlCompositeElement list, int level, A.DefaultRunProperties properties)
    {
        list.RemoveAllChildren();
        if (level == 1)
        {
            list.Append(new A.Level2ParagraphProperties(properties));
        }
        else
        {
            list.Append(new A.Level1ParagraphProperties(properties));
        }
    }

    private static P.Shape TemplatePlaceholder(uint id, uint index, P.PlaceholderValues type) => new(
        new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties { Id = id, Name = "Template" },
            new P.NonVisualShapeDrawingProperties(), new P.ApplicationNonVisualDrawingProperties(
                new P.PlaceholderShape { Type = type, Index = index })),
        new P.ShapeProperties(), new P.TextBody(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.EndParagraphRunProperties())));
}
