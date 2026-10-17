using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Helpers;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests;

public class ShapeFillTests : SCTest
{
    [Test]
    public void Fill_is_not_null()
    {
        // Arrange
        var pptx = TestAsset("021.pptx");
        var pres = new Presentation(pptx);
        var autoShape = pres.Slides[0].Shapes.First(sp => sp.Id == 108);

        // Act-Assert
        autoShape.Fill.Should().NotBeNull();
    }

    [Test]
    public void Picture_SetImage_updates_picture_fill()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var image = TestAsset("10 png image.png");
        var shape = new Presentation(pptx).Slides[2].Shapes.First(sp => sp.Id == 4);
        var fill = shape.Fill;
        var imageSizeBefore = fill.Picture!.AsByteArray().Length;

        // Act
        fill.Picture.Update(image);

        // Assert
        var imageSizeAfter = shape.Fill.Picture.AsByteArray().Length;
        imageSizeAfter.Should().NotBe(imageSizeBefore, "because image has been changed");
    }

    [Test]
    public void AutoShape_Fill_Type_returns_NoFill_When_shape_is_Not_filled()
    {
        // Arrange
        var autoShape = new Presentation(TestAsset("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 6);

        // Act
        var fillType = autoShape.Fill.Type;

        // Assert
        fillType.Should().Be(FillType.NoFill);
    }

    [Test]
    public void HexSolidColor_getter_returns_color_name()
    {
        // Arrange
        var autoShape = new Presentation(TestAsset("009_table.pptx")).Slides[1].Shapes.First(sp => sp.Id == 2);

        // Act
        var shapeSolidColorName = autoShape.Fill.Color;

        // Assert
        shapeSolidColorName.Should().BeEquivalentTo("ff0000");
    }

    [Test]
    public void Color_Getter_returns_color_hex()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shapeFill = pres.Slides[3].Shapes.First(sp => sp.Name == "Rectangle 3").Fill;

        // Act
        var colorHex = shapeFill.Color;

        // Assert
        colorHex.Should().BeEquivalentTo("FFAB40");
    }

    [Test]
    public void Alpha_returns_opacity_level_of_fill_color_in_percentages()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shapeFill = pres.Slides[3].Shapes.First(sp => sp.Name == "SolidSchemeAlpha").Fill;

        // Act
        var alpha = shapeFill.Alpha;

        // Assert
        alpha.Should().Be(60);
    }

    [Test]
    public void ThemeColorWithLuminanceLight_getter_returns_color_name()
    {
        // Arrange
        var autoShape = new Presentation(TestAsset("009_table.pptx")).Slides[3].Shapes
            .First(sp => sp.Name == "SolidSchemeLumLight");

        // Act
        var fill = autoShape.Fill;

        // Assert
        fill.LuminanceModulation.Should().Be(20);
        fill.LuminanceOffset.Should().Be(80);
    }

    [Test]
    public void Luminance_properties()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shapeFill = pres.Slides[3].Shapes.First(sp => sp.Name == "SolidSchemeLumDark").Fill;

        // Act-Assert
        shapeFill.LuminanceModulation.Should().Be(75);
        shapeFill.LuminanceOffset.Should().Be(0);
    }

    [Test]
    public void Picture_BinaryData_returns_binary_content_of_picture_image()
    {
        // Arrange
        var pptxStream = TestAsset("009_table.pptx");
        var pres = new Presentation(pptxStream);
        var shapeFill = pres.Slides[2].Shapes.Shape("AutoShape 1").Fill;

        // Act
        var imageBytes = shapeFill.Picture!.AsByteArray();

        // Assert
        imageBytes.Length.Should().BePositive();
    }

    [Test]
    public void SetColor_sets_green_color()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddShape(0, 0, 100, 100);
        var shape = slide.Shapes.Last();

        // Act
        shape.Fill!.SetColor("00FF00");

        // Assert
        ValidatePresentation(pres);
    }

    [Test]
    [TestCase("autoshape-case005_text-frame.pptx", 1, "AutoShape 1")]
    [TestCase("autoshape-case005_text-frame.pptx", 1, "AutoShape 2")]
    [TestCase("autoshape-grouping.pptx", 1, "AutoShape 1")]
    public void SetColor_sets_solid_color(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.Shape(shapeName);
        var shapeFill = shape.Fill;

        // Act
        shapeFill.SetColor("32a852");

        // Assert
        shapeFill.Color.Should().Be("32a852");
        ValidatePresentation(pres);
    }

    [Test]
    public void SetPicture_should_not_break_presentation()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shape = pres.Slide(2).Shape("AutoShape 2");
        var image = TestAsset("09 png image.png");
        var shapeFill = shape.Fill!;

        // Act
        shapeFill.SetPicture(image);

        // Assert
        ValidatePresentation(pres);
    }

    [TestCase("autoshape-case005_text-frame.pptx", 1, "AutoShape 1")]
    [TestCase("autoshape-case005_text-frame.pptx", 1, "AutoShape 2")]
    [TestCase("autoshape-grouping.pptx", 1, "AutoShape 1")]
    public void SetNoFill_removes_fill(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shapeFill = pres.Slides[slideNumber - 1].Shapes.Shape(shapeName).Fill;

        // Act
        shapeFill.SetNoFill();

        // Assert
        shapeFill.Type.Should().Be(FillType.NoFill);
        ValidatePresentation(pres);
    }

    [TestCase("009_table.pptx", 2, "AutoShape 1", FillType.Gradient)]
    [TestCase("009_table.pptx", 2, "AutoShape 2", FillType.Pattern)]
    [TestCase("009_table.pptx", 4, "Rectangle 3", FillType.Solid)]
    [TestCase("009_table.pptx", 3, "AutoShape 1", FillType.Picture)]
    [TestCase("autoshape-case003.pptx", 1, "AutoShape 1", FillType.SlideBackground)]
    public void SetNoFill_replaces_existing_fill_and_survives_reopening(
        string file, int slideNumber, string shapeName, FillType originalType)
    {
        // Arrange
        using var pres = new Presentation(TestAsset(file));
        var shape = pres.Slide(slideNumber).Shape(shapeName);
        var fill = shape.Fill;
        fill.Type.Should().Be(originalType);
        _ = fill.Picture;

        // Act
        fill.SetNoFill();
        fill.SetNoFill();

        // Assert
        fill.Type.Should().Be(FillType.NoFill);
        fill.Color.Should().BeNull();
        fill.Picture.Should().BeNull();
        using var reopened = SaveAndOpenPresentation(pres);
        reopened.Slide(slideNumber).Shape(shapeName).Fill.Type.Should().Be(FillType.NoFill);
        using var sdk = SaveAndOpenPresentationAsSdk(pres);
        var slideId = sdk.PresentationPart!.Presentation.SlideIdList!.Elements<P.SlideId>().ElementAt(slideNumber - 1);
        var slidePart = (SlidePart)sdk.PresentationPart.GetPartById(slideId.RelationshipId!);
        var sdkShape = slidePart.Slide.Descendants<P.Shape>()
            .Single(item => item.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name == shapeName);
        var properties = sdkShape.ShapeProperties!;
        properties.Elements<A.NoFill>().Should().ContainSingle();
        properties.Elements().Where(item => item is A.SolidFill or A.GradientFill or A.PatternFill
            or A.BlipFill or A.GroupFill).Should().BeEmpty();
        new OpenXmlValidator().Validate(properties).Select(error => error.Description).Should().BeEmpty();
    }

    [TestCase(false)]
    [TestCase(true)]
    public void SetNoFill_preserves_outline_and_effects_in_valid_xml_order(bool useGroupFill)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Shapes.AddShape(0, 0, 100, 100);
        var shape = pres.Slides[0].Shapes.Last();
        shape.Fill.SetColor("FF0000");
        shape.Outline.Weight = 2;
        using var stream = new MemoryStream();
        pres.Save(stream);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, true))
        {
            var shapeProperties = document.PresentationPart!.SlideParts.Single().Slide
                .Descendants<P.ShapeProperties>().Single();
            if (useGroupFill)
            {
                shapeProperties.RemoveAllChildren<A.SolidFill>();
                shapeProperties.AddChild(new A.GroupFill(), true);
            }

            shapeProperties.AddChild(new A.EffectList(), true);
        }

        stream.Position = 0;
        using var edited = new Presentation(stream);
        var editedShape = edited.Slides[0].Shapes.Single();

        // Act
        editedShape.Fill.SetNoFill();

        // Assert
        editedShape.Outline.Weight.Should().Be(2);
        using var sdk = SaveAndOpenPresentationAsSdk(edited);
        var properties = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<P.ShapeProperties>().Single();
        properties.Elements<A.NoFill>().Should().ContainSingle();
        properties.Elements<A.GroupFill>().Should().BeEmpty();
        properties.GetFirstChild<A.Outline>().Should().NotBeNull();
        properties.GetFirstChild<A.EffectList>().Should().NotBeNull();
        new OpenXmlValidator().Validate(properties).Select(error => error.Description).Should().BeEmpty();
    }

    [Test]
    public void SetNoFill_removes_fill_from_individual_grouped_shape()
    {
        // Arrange
        using var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var group = pres.Slide(1).Shape("Group 2");
        var shape = group.GroupedShapes.Shape("Shape 1");
        shape.Fill.SetColor("FF0000");
        var sibling = group.GroupedShapes.Shape("Shape 2");
        sibling.Fill.SetColor("00FF00");

        // Act
        shape.Fill.SetNoFill();

        // Assert
        shape.Fill.Type.Should().Be(FillType.NoFill);
        sibling.Fill.Color.Should().Be("00FF00");
        using var reopened = SaveAndOpenPresentation(pres);
        var reopenedGroup = reopened.Slide(1).Shape("Group 2");
        reopenedGroup.GroupedShapes.Shape("Shape 1").Fill.Type.Should().Be(FillType.NoFill);
        reopenedGroup.GroupedShapes.Shape("Shape 2").Fill.Color.Should().Be("00FF00");
    }

    [TestCase("table-case001.pptx", 1, "Table 1")]
    public void SetColor_sets_solid_color_as_fill_of_Table_Cell(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var table = pres.Slide(slideNumber).Shape(shapeName).Table;
        var shapeFill = table[0, 0].Fill;

        // Act
        shapeFill.SetColor("32a852");

        // Assert
        shapeFill.Color.Should().Be("32a852");
        ValidatePresentation(pres);
    }

    [Test]
    [TestCase("009_table.pptx", 2, "AutoShape 2")]
    public void SetFill_sets_NoFill_of_Table_Cell(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slides[slideNumber - 1].Shapes.Shape(shapeName);
        var shapeFill = shape.Fill;

        // Act
        shapeFill.SetNoFill();

        // Assert
        shape.Fill.Type.Should().Be(FillType.NoFill);
        ValidatePresentation(pres);
    }

    [Theory]
    [SlideShape("008.pptx", slideNumber: 1, shapeName: "AutoShape 1")]
    [SlideShape("autoshape-case009.pptx", slideNumber: 1, shapeName: "AutoShape 1")]
    [LayoutShape("autoshape-case003.pptx", slideLayoutNumber: 1, shapeName: "AutoShape 1")]
    [MasterShape("autoshape-case003.pptx", shapeName: "AutoShape 1")]
    public void SetPicture_updates_fill_with_specified_picture_image_When_shape_is_Not_filled(IShape shape)
    {
        // Arrange
        var fill = shape.Fill;
        var imageStream = TestAsset("09 png image.png");

        // Act
        fill.SetPicture(imageStream);

        // Assert
        var pictureBytes = fill.Picture!.AsByteArray();
        var imageBytes = imageStream.ToArray();
        pictureBytes.SequenceEqual(imageBytes).Should().BeTrue();
    }

    [Test]
    [SlideShape("009_table.pptx", 2, 6, FillType.NoFill)]
    [SlideShape("009_table.pptx", 2, 2, FillType.Solid)]
    [SlideShape("009_table.pptx", 3, 4, FillType.Picture)]
    [SlideShape("009_table.pptx", 2, "AutoShape 2", FillType.Pattern)]
    [SlideShape("autoshape-case003.pptx", 1, "AutoShape 1", FillType.SlideBackground)]
    public void Type_returns_fill_type(IShape shape, FillType expectedFill)
    {
        // Act
        var fillType = shape.Fill.Type;

        // Assert
        fillType.Should().Be(expectedFill);
    }

    [Test]
    public void Type_returns_Gradient_fill_type()
    {
        // Act
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shapeFill = pres.Slides[1].Shapes.Shape<IShape>("AutoShape 1").Fill;

        // Act-Assert
        shapeFill.Type.Should().Be(FillType.Gradient);
    }

    [Test]
    [TestCase("009_table.pptx", 2, "AutoShape 2")]
    public void SetColor_replaces_picture_with_solid_color(string file, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var shape = pres.Slide(slideNumber).Shapes.Shape(shapeName);
        var shapeFill = shape.Fill;
        var image = TestAsset("09 png image.png");
        var greenColor = "32a852";

        // Act
        shapeFill.SetPicture(image);
        shapeFill.SetColor(greenColor);

        // Assert
        shapeFill.Color.Should().Be(greenColor);
        ValidatePresentation(pres);
    }
}
