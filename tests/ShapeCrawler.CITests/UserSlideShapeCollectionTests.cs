using FluentAssertions;
using ImageMagick;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.CITests;

public class UserSlideShapeCollectionTests : SCTest
{
    [Test]
    public void AddPicture_sets_jpeg_mime()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("04 heic image.heic");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = shapes.Last().Picture;
        
        picture.Image!.Mime.Should().Be("image/jpeg");
        
        var actualImage = new MagickImage(picture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.jpg"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        ValidatePresentation(pres);
    }

    [TestCase("webp image.webp")]
    [TestCase("01 avif image.avif")]
    [TestCase("02 bmp image.bmp")]
    public void AddPicture_sets_png_mime(string image)
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var shapes = pres.Slides[0].Shapes;
        var imageStream = TestAsset(image);

        // Act
        shapes.AddPicture(imageStream);

        // Assert
        var addedPicture = shapes.Last().Picture;
        
        addedPicture.Image!.Mime.Should().Be("image/png");
        
        var actualImage = new MagickImage(addedPicture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.png"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        ValidatePresentation(pres);
    }
    
    [Test]
    public void AddPicture_adds_picture_from_ico_image()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("05 ico image.ico");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = shapes.Last().Picture;
        picture.Image!.Mime.Should().Be("image/png");
        var actualImage = new MagickImage(picture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference ico.png"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        ValidatePresentation(pres);
    }
    
    [Test]
    public void AddAudio_adds_audio_shape_with_MP3_content()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var mp3 = TestAsset("064 mp3.mp3");
        var pres = new Presentation(pptx);
        var shapes = pres.Slides[1].Shapes;
        int xPtCoordinate = 225;
        int yPtCoordinate = 75;

        // Act
        shapes.AddAudio(xPtCoordinate, yPtCoordinate, mp3);

        pres.Save();
        pres = new Presentation(pptx);
        var addedAudio = pres.Slides[1].Shapes.Last();

        // Assert
        addedAudio.X.Should().Be(xPtCoordinate);
        addedAudio.Y.Should().Be(yPtCoordinate);
    }
    
    [Test]
    public void AddSmartArt_adds_Basic_Block_List_SmartArt_graphic()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slide(1);
        const int x = 50;
        const int y = 50;
        const int width = 400;
        const int height = 300;
        
        // Act
        var smartArtShape = slide.Shapes.AddSmartArt(x, y, width, height, SmartArtType.BasicBlockList);
        
        // Assert
        ValidatePresentation(pres);
        smartArtShape.SmartArt.Should().NotBeNull();
        smartArtShape.X.Should().Be(x);
        smartArtShape.Y.Should().Be(y);
        smartArtShape.Width.Should().Be(width);
        smartArtShape.Height.Should().Be(height);
        var slidePart = pres.GetSdkPresentationDocument().PresentationPart!.SlideParts.First();
        var relationshipTypes = slidePart.Parts.Select(part => part.OpenXmlPart.RelationshipType).ToArray();
        relationshipTypes.Should().Contain("http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData");
    }
}
