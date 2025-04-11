using FluentAssertions;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Load;

public class ShapeCollectionTests : SCTest
{
    [Test]
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
        Task.Delay(1000).Wait();
        shapes = loadedPres.Slides[0].Shapes;
        image.Position = 0;
        shapes.AddPicture(image);

        // Assert
        var presDocument = SaveAndOpenPresentationAsSdk(loadedPres);
        var imageParts = presDocument.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(2,
            "SVG image adds two parts: One for the vector and one for the auto-generated raster");
        loadedPres.Validate();
    }
    
    [Test]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_svg_image_is_added_twice()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var svgImage = TestAsset("063 vector image.svg");

        // Act
        shapes.AddPicture(svgImage);
        Task.Delay(1000).Wait();
        svgImage.Position = 0;
        shapes.AddPicture(svgImage);

        // Assert
        var checkXml = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = checkXml.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).ToArray();
        imageParts.Length.Should().Be(2,
            "SVG image adds two parts: One for the vector and one for the auto-generated raster");
    }
    
    [TestCase("08 jpeg image-500w.jpg")]
    [TestCase("09 png image.png")]
    [TestCase("03 gif image.gif")]
    [TestCase("07 tiff image.tiff")]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_image_is_added_a_second_apart(string fileName)
    {
        // Arrange
        var pres = new Presentation();
        pres.Slides.AddEmptySlide(SlideLayoutType.Blank);
        var shapesSlide1 = pres.Slides[0].Shapes;
        var shapesSlide2 = pres.Slides[1].Shapes;

        var image = TestAsset(fileName);

        // Act
        shapesSlide1.AddPicture(image);
        Task.Delay(1000).Wait();
        shapesSlide2.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(1);
    }
}