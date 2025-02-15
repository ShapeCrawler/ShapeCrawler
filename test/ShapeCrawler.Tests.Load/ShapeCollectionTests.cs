using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Load;

public class ShapeCollectionTests : SCTest
{
    [Test]
    [Repeat(40)]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_svg_image_is_added_to_a_loaded_presentation()
    {
        // Arrange
        var pres = new Presentation();
        pres.Slides.AddEmptySlide(SlideLayoutType.Blank);
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("063 vector image.svg");
        shapes.AddPicture(image);
        var loadedPres = SaveAndOpenPresentation(pres);

        // Act
        shapes = loadedPres.Slides[0].Shapes;
        shapes.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(loadedPres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(2,
            "SVG image adds two parts: One for the vector and one for the auto-generated raster");
        loadedPres.Validate();
    }
    
    [Test]
    [Repeat(40)]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_svg_image_is_added_twice()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var svgImage = TestAsset("063 vector image.svg");

        // Act
        shapes.AddPicture(svgImage);
        shapes.AddPicture(svgImage);

        // Assert
        var checkXml = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = checkXml.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).ToArray();
        imageParts.Length.Should().Be(2,
            "SVG image adds two parts: One for the vector and one for the auto-generated raster");
    }
}