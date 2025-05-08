using System.Diagnostics.CodeAnalysis;
using System.Xml;
using FluentAssertions;
using ImageMagick;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Groups;


// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.DevTests;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
public class ShapeCollectionTests : SCTest
{
    [Test]
    public void Add_adds_shape()
    {
        // Arrange
        var pres = new Presentation(TestAsset("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.Shape("TextBox")!;
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        shapes.Shape("TextBox 2").Should().NotBeNull();
    }

    [Test]
    public void Add_adds_table()
    {
        // Arrange
        var pres = new Presentation(TestAsset("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.Shape("Table 1")!;
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        var addedShape = shapes.Last();
        addedShape.Should().BeAssignableTo<ITable>();
    }

    [Test]
    public void Add_adds_picture_from_same_slide()
    {
        // Arrange
        var pres = new Presentation(TestAsset("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.Shape("Picture")!;
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        shapes.Shape("Picture 1").Should().NotBeNull();
        pres.Validate();
    }

    [Test]
    [Explicit("Failing test for https://github.com/ShapeCrawler/ShapeCrawler/issues/935")]   
    public void Add_adds_picture_to_another_slide()
    {
        // Arrange
        var pres = new Presentation(TestAsset("053_add_shapes.pptx"));
        var copyingShape = pres.Slides[0].Shapes.Shape("Picture")!;
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.Add(copyingShape);

        // Assert
        shapes.Shape("Picture 1").Should().NotBeNull();
        pres.Validate();
    }

    [Test]
    public void Contains_expected_count_of_each_shape_type()
    {
        // Arrange
        var pres = new Presentation(TestAsset("003.pptx"));
        var shapes = pres.Slides.First().Shapes;

        // Act & Assert
        shapes.Count(sp => sp.ShapeContent == ShapeContent.Chart).Should().Be(1);
        shapes.Count(sp => sp.ShapeContent == ShapeContent.Picture).Should().Be(1);
        shapes.Count(sp => sp.ShapeContent == ShapeContent.Table).Should().Be(1);
        shapes.Count(sp => sp.ShapeContent == ShapeContent.Group).Should().Be(1);
    }

    [Test]
    public void Contains_picture()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var shape = pres.Slide(2).Shapes.First(sp => sp.Id == 3);

        // Act-Assert
        var picture = shape as IPicture;
        picture.Should().NotBeNull();
    }

    [Test]
    public void Contains_Media_Shape()
    {
        // Arrange
        var pptxStream = TestAsset("audio-case001.pptx");
        var pres = new Presentation(pptxStream);
        IShape shape = pres.Slides[0].Shapes.First(sp => sp.Id == 8);

        // Act
        bool isMediaShape = shape is IMediaShape;

        // Assert
        isMediaShape.Should().BeTrue();
    }

    [Test]
    public void Contains_Connection_shape()
    {
        var pres = new Presentation(TestAsset("001.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act-Assert
        shapes.Should().Contain(shape => shape.Id == 10 && shape is ILine && shape.GeometryType == Geometry.Line);
    }

    [Test]
    public void Contains_Video_shape()
    {
        // Arrange
        var pptx = TestAsset("040_video.pptx");
        var pres = new Presentation(pptx);
        IShape shape = pres.Slides[0].Shapes.First(sp => sp.Id == 8);

        // Act
        bool isVideo = shape is IMediaShape;

        // Act-Assert
        isVideo.Should().BeTrue();
    }

    [Test]
    public void AddLine_adds_a_new_Line_shape_from_raw_open_xml_content()
    {
        // Arrange
        var pres = new Presentation();
        var xml = StringOf("line-shape.xml");
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(xml);

        // Assert
        var addedLine = shapes.Last();
        addedLine.Id.Should().Be(1);
        shapes.Count.Should().Be(1);
    }

    [Test]
    public void AddLine_adds_line_Right_Up()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 10, startPointY: 10, endPointX: 20, endPointY: 5);

        // Assert
        var addedLine = (ILine)shapes.Last();
        shapes.Should().ContainSingle();
        addedLine.ShapeContent.Should().Be(ShapeContent.Line);
        addedLine.StartPoint.X.Should().Be(10);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(20);
        addedLine.EndPoint.Y.Should().Be(5);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Up_Up()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 10, startPointY: 10, endPointX: 10, endPointY: 5);

        // Assert
        var addedLine = (ILine)shapes.Last();
        addedLine.StartPoint.X.Should().Be(10);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(10);
        addedLine.EndPoint.Y.Should().Be(5);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Left_Up()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 100, startPointY: 50, endPointX: 40, endPointY: 20);

        // Assert
        var addedLine = (ILine)shapes.Last();
        addedLine.StartPoint.X.Should().Be(100);
        addedLine.StartPoint.Y.Should().Be(50);
        addedLine.EndPoint.X.Should().Be(40);
        addedLine.EndPoint.Y.Should().Be(20);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Left_Down()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 10, endPointX: 40, endPointY: 20);

        // Assert
        var addedLine = (ILine)shapes.Last();
        addedLine.StartPoint.X.Should().Be(50);
        addedLine.StartPoint.Y.Should().Be(10);
        addedLine.EndPoint.X.Should().Be(40);
        addedLine.EndPoint.Y.Should().Be(20);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Right_Right()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Assert
        var line = (ILine)shapes.Last();
        line.StartPoint.X.Should().Be(50);
        line.StartPoint.Y.Should().Be(60);
        line.EndPoint.X.Should().Be(100);
        line.EndPoint.Y.Should().Be(60);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 50, startPointY: 60, endPointX: 100, endPointY: 60);

        // Assert
        shapes.Should().ContainSingle();
        var line = (ILine)shapes.Last();
        line.ShapeContent.Should().Be(ShapeContent.Line);
        line.X.Should().Be(50);
        line.Y.Should().Be(60);
        pres.Validate();
    }

    [Test]
    public void AddLine_adds_line_Left_Left()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddLine(startPointX: 100, startPointY: 50, endPointX: 80, endPointY: 50);

        // Assert
        var line = (ILine)shapes.Last();
        line.StartPoint.X.Should().Be(100);
        line.StartPoint.Y.Should().Be(50);
        line.EndPoint.X.Should().Be(80);
        line.EndPoint.Y.Should().Be(50);
        pres.Validate();
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
        var addedAudio = pres.Slides[1].Shapes.OfType<IMediaShape>().Last();

        // Assert
        addedAudio.X.Should().Be(xPtCoordinate);
        addedAudio.Y.Should().Be(yPtCoordinate);
    }

    [Test]
    public void AddAudio_adds_audio_shape_with_WAVE_content()
    {
        // Arrange
        var wav = TestAsset("071 wav.wav");
        var pres = new Presentation(TestAsset("001.pptx"));
        var shapes = pres.Slides[1].Shapes;

        // Act
        shapes.AddAudio(300, 100, wav, AudioType.WAVE);

        // Assert
        var addedAudio = pres.Slides[1].Shapes.OfType<IMediaShape>().Last();
        addedAudio.X.Should().Be(300);
    }
    
    [Test, Explicit("Should be implemented with https://github.com/ShapeCrawler/ShapeCrawler/issues/581")]
    public void AddAudio_adds_audio_shape_with_the_default_start_mode_In_Click_Sequence()
    {
        // Arrange
        var pres = new Presentation();
        var mp3 = TestAsset("064 mp3.mp3");
        var shapes = pres.Slide(1).Shapes;

        // Act
        shapes.AddAudio(x: 300, y: 100, mp3, AudioType.MP3);

        // Assert
        pres = SaveAndOpenPresentation(pres);
        var addedAudio = pres.Slide(1).First<IMediaShape>();
        pres.Validate();
        addedAudio.StartMode.Should().Be(AudioStartMode.InClickSequence);
    }
    
    [Test]
    public void AddVideo_adds_Video_shape()
    {
        // Arrange
        var preStream = TestAsset("001.pptx");
        var presentation = new Presentation(preStream);
        var shapesCollection = presentation.Slides[1].Shapes;
        var videoStream = TestAsset("test-video.mp4");
        int xPxCoordinate = 300;
        int yPxCoordinate = 100;

        // Act
        shapesCollection.AddVideo(xPxCoordinate, yPxCoordinate, videoStream);

        // Assert
        presentation.Save();
        presentation = new Presentation(preStream);
        var addedVideo = presentation.Slides[1].Shapes.OfType<IMediaShape>().Last();
        addedVideo.X.Should().Be(xPxCoordinate);
        addedVideo.Y.Should().Be(yPxCoordinate);
    }
    
    [Test]
    public void AddPicture_adds_svg_picture()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("063 vector image.svg");
        image.Position = 0;

        // Act
        shapes.AddPicture(image);

        // Assert
        shapes.Should().HaveCount(1);
        var picture = (IPicture)shapes.Last();
        picture.ShapeContent.Should().Be(ShapeContent.Picture);
        picture.Height.Should().Be(75);
        picture.Width.Should().Be(75);
        pres.Validate();
    }
    
    [Test]
    [Category("issue-883")]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_svg_image_is_added_on_two_different_slides()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(1).SlideLayout("Blank");
        pres.Slides.Add(layout.Number);
        var shapesSlide1 = pres.Slides[0].Shapes;
        var shapesSlide2 = pres.Slides[1].Shapes;
        var image = TestAsset("063 vector image.svg");

        // Act
        shapesSlide1.AddPicture(image);
        image.Position = 0;
        shapesSlide2.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(2,
            "SVG image adds two parts: One for the vector and one for the auto-generated raster");
    }

    [Test]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_image_is_added_on_two_different_slides()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMasters[0].SlideLayouts.First(l => l.Name == "Blank");
        pres.Slides.Add(layout.Number);
        var shapesSlide1 = pres.Slides[0].Shapes;
        var shapesSlide2 = pres.Slides[1].Shapes;

        var image = TestAsset("09 png image.png");

        // Act
        shapesSlide1.AddPicture(image);
        shapesSlide2.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(1);
    }
    
    [Test]
    [Explicit("Should be fixed")]
    public void AddPicture_should_not_duplicate_the_image_source_When_slide_is_copied()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMaster(0).SlideLayouts.First(l => l.Name == "Blank");
        pres.Slides.Add(layout.Number);
        var slide = pres.Slides[0];
        var shapes = slide.Shapes;
        var image = TestAsset("09 png image.png");
        shapes.AddPicture(image);

        // Act
        pres.Slides.Add(slide);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(1);
    }

    [Test]
    public void AddPicture_sets_valid_svg_content()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("063 vector image.svg");
        image.Position = 0;
        shapes.AddPicture(image);
        var picture = (IPicture)shapes.Last();

        // Act
        var svgContent = picture.SvgContent;
        
        // Assert
        svgContent.Should().Contain("<svg");
    }
    
    [Test]
    public void AddPicture_preserves_original_1600_width()
    {
        // Arrange
        var pres = new Presentation();
        var imageStream = TestAsset("11 image 1600x690.jpg");
        var shapes = pres.Slide(1).Shapes;
        
        // Act
        shapes.AddPicture(imageStream);
        
        // Assert
        var addedPicture = shapes.Last<IPicture>();
        var image = new MagickImage(addedPicture.Image!.AsByteArray());
        image.Width.Should().Be(1600);
    }
    
    [Test]
    public void AddPicture_svg_with_text_matches_reference()
    {
        // ARRANGE

        // This presentation contains the same SVG we're adding below, manually
        // dragged in while running PowerPoint
        var pres = new Presentation(TestAsset("055_svg_with_text.pptx"));
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("066 1x1.svg");
        image.Position = 0;

        // ACT
        shapes.AddPicture(image);

        // ASSERT
        var picture = (IPicture)shapes.First(shape => shape.Name.StartsWith("Picture"));
        var xml = new XmlDocument { PreserveWhitespace = true };
        xml.LoadXml(picture.SvgContent);
        var textTagRandomChild = xml.GetElementsByTagName("text").OfType<XmlElement>().First().ChildNodes.Item(0);
        textTagRandomChild.Should().NotBeOfType<XmlSignificantWhitespace>("Text tags must not contain whitespace");
        
        // The above assertion does guard against the root cause of the bug 
        // which led to this test. However, the true test comes from loading
        // these up in PowerPoint and ensure the added image looks like the
        // existing image.
        pres.Validate();
    }

    [Test]
    public void AddPicture_too_large_adds_picture()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("png image-large.png");
        image.Position = 0;

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = (IPicture)shapes.Last();

        // These values are reasonable range for size of an added image
        picture.Height.Should().BeGreaterThan(0);
        picture.Height.Should().BeLessThan(2400);
        picture.Width.Should().BeGreaterThan(0);
        picture.Width.Should().BeLessThan(2400);

        // Ensure aspect ratio has been maintained
        var aspect = picture.Width / picture.Height;
        aspect.Should().Be(100);

        pres.Validate();
    }

    [Test]
    public void AddPicture_adds_svg_picture_no_width_height_tags()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("070 vector image-wide.svg");
        image.Position = 0;

        // Act
        shapes.AddPicture(image);

        // Assert
        // These values are the viewbox size of the test image, which is what
        // we'll be using since the image has no width or height tags
        var picture = (IPicture)shapes.Last();
        picture.Height.Should().BeApproximately(67.5m, 0.1m);
        picture.Width.Should().Be(210);
        pres.Validate();
    }
    
    [Test]
    public void AddPicture_adds_picture_with_transparent_background()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("067 vector image-blank.svg");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = (IPicture)shapes.Last();
        var imageMagickImage = new MagickImage(picture.Image!.AsByteArray());
        var pixels = imageMagickImage.GetPixels();
        pixels.Should().NotBeEmpty().And.AllSatisfy(x =>
        {
            x.ToColor().Should().Be(MagickColors.Transparent);
        });
    }

    [TestCase("09 png image.png", "image/png")]
    [TestCase("06 jpeg image.jpg", "image/jpeg")]
    [TestCase("03 gif image.gif", "image/gif")]
    [TestCase("07 tiff image.tiff", "image/tiff")]
    public void AddPicture_adds_should_set_valid_mime(string image, string expectedMime)
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var imageStream = TestAsset(image);

        // Act
        shapes.AddPicture(imageStream);

        // Assert
        shapes.Should().HaveCount(1);
        var picture = (IPicture)shapes.Last();
        picture.Image!.Mime.Should().Be(expectedMime);
        pres.Validate();
    }
    
    [TestCase("webp image.webp")]
    [TestCase("01 avif image.avif")]
    [TestCase("02 bmp image.bmp")]
    public void AddPicture_should_set_png_mime(string image)
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var imageStream = TestAsset(image);

        // Act
        shapes.AddPicture(imageStream);

        // Assert
        var addedPicture = shapes.Last<IPicture>();
        
        addedPicture.Image!.Mime.Should().Be("image/png");
        
        var actualImage = new MagickImage(addedPicture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.png"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        pres.Validate();
    }
    
    [Test]
    [Explicit("Should be fixed with https://github.com/ShapeCrawler/ShapeCrawler/issues/892")]
    public void AddPicture_adds_picture_from_ico_image()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("ico image.ico");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = (IPicture)shapes.Last();
        picture.Image!.Mime.Should().Be("image/png");
        var actualImage = new MagickImage(picture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.png"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        pres.Validate();
    }
    
    [Test]
    public void AddPicture_adds_should_set_jpeg_mime()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("04 heic image.heic");

        // Act
        shapes.AddPicture(image);

        // Assert
        var picture = shapes.Last<IPicture>();
        
        picture.Image!.Mime.Should().Be("image/jpeg");
        
        var actualImage = new MagickImage(picture.Image!.AsByteArray());
        var expectedImage = new MagickImage(TestAsset("reference image.jpg"));
        actualImage.GetPixels().Should().BeEquivalentTo(expectedImage.GetPixels());
        
        pres.Validate();
    }

    [Test]
    public void AddPicture_throws_exception_when_the_specified_stream_is_non_image()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        var stream = TestAsset("autoshape-case011_save-as-png.pptx");

        // Act
        var addingPicture = () => shapes.AddPicture(stream);

        // Assert
        addingPicture.Should().Throw<SCException>();
    }
    
    [Test]
    public void AddPicture_adds_picture_with_correct_Height()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("09 png image.png");

        // Act
        shapes.AddPicture(image);

        // Assert
        var addedPicture = shapes.Last();
        addedPicture.Height.Should().Be(225);
    }

    [Test]
    public void AddPicture_adds_picture_with_correct_mime()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("06 jpeg image.jpg");

        // Act
        shapes.AddPicture(image);

        // Assert
        var addedPictureImage = shapes.Last<IPicture>().Image!;
        addedPictureImage.Mime.Should().Be("image/jpeg");
    }
    
    [TestCase("08 jpeg image-500w.jpg")]
    [TestCase("09 png image.png")]
    [TestCase("03 gif image.gif")]
    [TestCase("07 tiff image.tiff")]
    public void AddPicture_should_not_change_the_underlying_file_size(string image)
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var imageStream = TestAsset(image);
        const int fileSizeTolerance = 2000;

        // Act
        shapes.AddPicture(imageStream);

        // Assert
        var addedPictureImage = shapes.Last<IPicture>().Image!;
        addedPictureImage.AsByteArray().Length.Should().BeLessOrEqualTo((int)imageStream.Length + fileSizeTolerance);
    }
    
    [Test]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_image_is_added_to_a_loaded_presentation()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMasters[0].SlideLayouts.First(l => l.Name == "Blank");
        pres.Slides.Add(layout.Number);
        var shapes = pres.Slides[0].Shapes;
        var image = TestAsset("09 png image.png");

        // Act
        shapes.AddPicture(image);
        var presLoaded = SaveAndOpenPresentation(pres);
        shapes = presLoaded.Slides[1].Shapes;
        shapes.AddPicture(image);

        // Assert
        var sdkPres = SaveAndOpenPresentationAsSdk(presLoaded);
        var imageParts = sdkPres.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).Select(imagePart => imagePart.Uri)
            .ToHashSet();
        imageParts.Count.Should().Be(1);
    }
    
    [Test]
    public void AddShape_adds_rectangle_with_valid_id_and_name()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case011_save-as-png.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70);

        // Assert
        var autoShape = shapes.Last();
        autoShape.Name.Should().Be("Rectangle");
        autoShape.Id.Should().Be(7);
        pres.Validate();
    }

    [Test]
    public void AddRectangle_adds_Rectangle_in_the_New_Presentation()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70);

        // Assert
        var rectangle = shapes.Last();
        rectangle.GeometryType.Should().Be(Geometry.Rectangle);
        rectangle.X.Should().Be(50);
        rectangle.Y.Should().Be(60);
        rectangle.Width.Should().Be(100);
        rectangle.Height.Should().Be(70);
        rectangle.TextBox!.Paragraphs.Count.Should().Be(1);
        rectangle.Outline.HexColor.Should().BeNull();
        pres.Validate();
    }

    [Test]
    public void AddShape_adds_Rounded_Rectangle()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70, Geometry.RoundedRectangle);

        // Assert
        var roundedRectangle = shapes.Last();
        roundedRectangle.GeometryType.Should().Be(Geometry.RoundedRectangle);
        roundedRectangle.Name.Should().Be("RoundedRectangle");
        roundedRectangle.Outline.HexColor.Should().BeNull();
        pres.Validate();
    }

    [Test]
    public void AddShape_adds_Top_Corners_Rounded_Rectangle()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddShape(50, 60, 100, 70, Geometry.TopCornersRoundedRectangle);

        // Assert
        var addedTopCornersRoundedRectangle = shapes.Last();
        addedTopCornersRoundedRectangle.GeometryType.Should().Be(Geometry.TopCornersRoundedRectangle);
        addedTopCornersRoundedRectangle.Name.Should().Be("TopCornersRoundedRectangle");
        pres.Validate();
    }

    [Test]
    public void AddTable_adds_table()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;

        // Act
        shapes.AddTable(x: 50, y: 60, columnsCount: 3, rowsCount: 2);

        // Assert
        var table = (ITable)shapes.Last();
        table.Columns.Should().HaveCount(3);
        table.Rows.Should().HaveCount(2);
        table.Id.Should().Be(1);
        table.Name.Should().Be("Table 1");
        table.Columns[0].Width.Should().BeApproximately(213.33m, 0.01m);
        pres.Validate();
    }
    
    [Test]
    [LayoutShape("autoshape-case004_subtitle.pptx", 1, "Group 1")]
    [MasterShape("autoshape-case004_subtitle.pptx", "Group 1")]
    public void GetByName_returns_shape_by_specified_name(IShape shape)
    {
        // Arrange
        var groupShape = (IGroup)shape;
        var shapeCollection = groupShape.Shapes;
            
        // Act
        var resultShape = shapeCollection.Shape<IShape>("AutoShape 1");

        // Assert
        resultShape.Should().NotBeNull();
    }
    
    [Test]
    [TestCase("002.pptx", 1,4)]
    [TestCase("003.pptx", 1,5)]
    [TestCase("013.pptx", 1,4)]
    [TestCase("023.pptx", 1,1)]
    [TestCase("014.pptx", 3,5)]
    [TestCase("009_table.pptx", 1, 6)]
    [TestCase("009_table.pptx", 2, 8)]
    public void Count_returns_number_of_shapes(string pptxName, int slideNumber, int expectedCount)
    {
        // Arrange
        var pres = new Presentation(TestAsset(pptxName));
        var slide = pres.Slides[slideNumber - 1];

        // Act
        int shapesCount = slide.Shapes.Count;

        // Assert
        shapesCount.Should().Be(expectedCount);
    }
    
    [Test]
    public void Count_returns_one_When_presentation_contains_one_slide()
    {
        // Act
        var pptx17 = TestAsset("017.pptx");
        var pres17 = new Presentation(pptx17);        
        var pptx16 = TestAsset("016.pptx");
        var pres16 = new Presentation(pptx16);
        var numberSlidesCase1 = pres17.Slides.Count;
        var numberSlidesCase2 = pres16.Slides.Count;

        // Assert
        numberSlidesCase1.Should().Be(1);
        numberSlidesCase2.Should().Be(1);
    }
    
    [Test]
    public void Add_adds_slide_from_the_Same_presentation()
    {
        // Arrange
        var pptxStream = TestAsset("003 chart.pptx");
        var pres = new Presentation(pptxStream);
        var expectedSlidesCount = pres.Slides.Count + 1;
        var slideCollection = pres.Slides;
        var addingSlide = slideCollection[0];

        // Act
        pres.Slides.Add(addingSlide);

        // Assert
        pres.Slides.Count.Should().Be(expectedSlidesCount);
    }
    
    [Test]
    public void Add_adds_slide_After_updating_chart_series()
    {
        // Arrange
        var pptx = TestAsset("001 bar chart.pptx");
        var pres = new Presentation(pptx);
        var chart = pres.Slides[0].Shapes.Shape<IChart>("Bar Chart 1");
        var expectedSlidesCount = pres.Slides.Count + 1;

        // Act
        chart.SeriesCollection[0].Points[0].Value = 1;
        pres.Slides.Add(pres.Slides[0]);
        
        // Assert
        pres.Slides.Count.Should().Be(expectedSlidesCount);
    }

    [Test]
    public void AddEmptySlide_adds_New_slide()
    {
        // Arrange
        var pptx = TestAsset("autoshape-grouping.pptx");
        var pres = new Presentation(pptx);
        var layout = pres.SlideMasters[0].SlideLayouts[0]; 
        var slides = pres.Slides;

        // Act
        slides.Add(layout.Number);

        // Assert
        var addedSlide = slides.Last();
        addedSlide.Should().NotBeNull();
        pres.Validate();
    }
    
    [Test]
    public void Add_adds_a_new_slide()
    {
        // Arrange
        var pres = new Presentation();
        var layout = pres.SlideMasters[0].SlideLayouts.First(l => l.Name == "Blank");
        var slides = pres.Slides;

        // Act
        slides.Add(layout.Number);

        // Assert
        slides[1].Shapes.Should().HaveCount(0);
    }

    [Test]
    public void Add_adds_slide()
    {
        // Arrange
        var pres = new Presentation(TestAsset("017.pptx"));
        var layout = pres.SlideMaster(1).SlideLayout("Title and Content");

        // Act
        pres.Slides.Add(layout.Number);

        // Assert
        pres.Slide(2).Shapes.Count.Should().Be(2);
    }
    
    [Test]
    public void Remove()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes; 
        shapes.AddShape(10, 10, 10, 10);

        // Act
        shapes.Remove(shapes.Last());
        
        // Assert
        shapes.Should().HaveCount(0);
        pres.Validate();
    }
    
    [Test]
    public void AddPicture_should_not_duplicate_the_image_source_When_the_same_png_image_is_added_twice()
    {
        // Arrange
        var pres = new Presentation(TestAsset("008.pptx"));
        var shapes = pres.Slide(1).Shapes;
        var pngImage = TestAsset("09 png image.png");

        // Act
        shapes.AddPicture(pngImage);
        shapes.AddPicture(pngImage);

        // Assert
        var checkXml = SaveAndOpenPresentationAsSdk(pres);
        var imageParts = checkXml.PresentationPart!.SlideParts.SelectMany(slidePart => slidePart.ImageParts).ToArray();
        imageParts.Length.Should().Be(1);
    }
    
    [Test]
    public void AddPicture_sets_384_ppi_resolution_for_svg_picture()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        var svg = TestAsset("063 vector image.svg");
        
        // Act
        shapes.AddPicture(svg);
        
        // Assert
        var addedPicture = shapes.Last<IPicture>();
        var addedPictureInfo = new MagickImageInfo(addedPicture.Image!.AsByteArray());
        var addedPictureResolution = addedPictureInfo.Density!.ChangeUnits(DensityUnit.PixelsPerInch);
        addedPictureResolution.X.Should().BeApproximately(384, 0.1);
        addedPictureResolution.Y.Should().BeApproximately(384, 0.1);
    }
    
    [Test]
    public void AddPieChart_adds_pie_chart()
    {
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        var categoryValues = new Dictionary<string, double>{ { "1st Qtr", 10 }, { "2nd Qtr", 20 }, { "3rd Qtr", 30 } };
        
        // Act
        shapes.AddPieChart(100, 100, 400, 300, categoryValues, "Sales");
        
        // Assert
        shapes.Should().Contain(shape=> shape is IChart);
        pres.Validate();
    }

    [Test]
    public void AddBarChart_adds_bar_chart()
    {
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var categoryValues = new Dictionary<string, double>
        {
            { "Category 1", 10 },
            { "Category 2", 25 },
            { "Category 3", 15 }
        };
        string seriesName = "Sample Series";
        
        // Act
        shapes.AddBarChart(x, y, width, height, categoryValues, seriesName);
        
        // Assert
        shapes.Should().Contain(shape=> shape is IChart);
        pres.Validate();
    }
    
    [Test]
    public void AddScatterChart_adds_scatter_chart()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var pointValues = new Dictionary<double, double>
        {
            { 1.0, 5.2 },
            { 2.0, 7.3 },
            { 3.0, 8.1 },
            { 4.0, 9.5 },
            { 5.0, 12.3 }
        };
        string seriesName = "Data Series";
        
        // Act
        shapes.AddScatterChart(x, y, width, height, pointValues, seriesName);
        
        // Assert
        var chart = shapes.OfType<IChart>().Last();
        chart.Type.Should().Be(ChartType.ScatterChart);
        pres.Validate();
    }

    [Test]
    public void AddStackedColumnChart_adds_stacked_column_chart()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        int x = 100;
        int y = 100;
        int width = 500;
        int height = 300;
        var categoryValues = new Dictionary<string, IList<double>>
        {
            { "Category 1", new List<double> { 10, 20 } },
            { "Category 2", new List<double> { 30, 40 } },
            { "Category 3", new List<double> { 50, 60 } }
        };
        var seriesNames = new List<string> { "Series 1", "Series 2" };

        // Act
        shapes.AddStackedColumnChart(x, y, width, height, categoryValues, seriesNames);

        // Assert
        var chart = shapes.OfType<IChart>().Last();
        chart.Type.Should().Be(ChartType.BarChart);
        pres.Validate();
    }

    [Test]
    public void Group_groups_shapes()
    {
        // Arrange
        var pres = new Presentation();
        var shapes = pres.Slide(1).Shapes;
        shapes.AddShape(100, 100, 100, 100, Geometry.Rectangle, "Shape 1");
        shapes.AddShape(100, 200, 100, 100, Geometry.Rectangle, "Shape 2");
        var shape1 = shapes[0];
        var shape2 = shapes[1];
    
        // Act
        var group = shapes.Group([shape1, shape2]);
        
        // Assert
        group.Should().BeAssignableTo<IGroup>();
        group.Shapes.Should().HaveCount(2);
        pres.Validate();
    }
}