using FluentAssertions;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.CITests;

public class ShapeFillTests : SCTest
{
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