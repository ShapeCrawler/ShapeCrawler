using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tests;

public partial class UserSlideShapeCollectionTests
{
    [TestCase(AudioType.MP3, 0)]
    [TestCase(AudioType.MP3, 1)]
    [TestCase(AudioType.MP3, 2)]
    [TestCase(AudioType.WAVE, 0)]
    [TestCase(AudioType.WAVE, 1)]
    [TestCase(AudioType.WAVE, 2)]
    public void Add_copies_audio_with_its_content(AudioType audioType, int destination)
    {
        // Arrange
        var source = new Presentation(p => p.Slide());
        source.Slides.Add(source.SlideMaster(1).SlideLayout("Blank").Number);
        var sourceShapes = source.Slides[0].Shapes;
        using var audio = TestAsset(audioType == AudioType.MP3 ? "064 mp3.mp3" : "071 wav.wav");
        sourceShapes.AddAudio(125, 75, audio, audioType);
        var original = sourceShapes.Last();
        var expectedBytes = original.Media!.AsByteArray();
        var expectedMime = original.Media.MIME;
        using var sourceDocument = source.GetSdkPresentationDocument();
        var originalImageId = original.SdkOpenXmlElement.Descendants<A.Blip>().Single().Embed!.Value!;
        var expectedImage = CopyTestPartBytes(CopyTestSlidePart(sourceDocument, 1).GetPartById(originalImageId));
        var target = destination == 2 ? new Presentation(p => p.Slide()) : source;
        var targetSlide = target.Slides[destination == 1 ? 1 : 0];
        var shapeCount = targetSlide.Shapes.Count;

        // Act
        targetSlide.Shapes.Add(original);

        // Assert
        using var reopened = SaveAndOpenPresentation(target);
        var copiedShapes = reopened.Slide(targetSlide.Number).Shapes;
        var copied = copiedShapes.Last();
        copiedShapes.Should().HaveCount(shapeCount + 1);
        copied.ContentType.Should().Be(ShapeContentType.Audio);
        copied.X.Should().Be(original.X);
        copied.Y.Should().Be(original.Y);
        copied.Media!.MIME.Should().Be(expectedMime);
        copied.Media.AsByteArray().Should().Equal(expectedBytes);
        original.Media.AsByteArray().Should().Equal(expectedBytes);
        using var reopenedDocument = reopened.GetSdkPresentationDocument();
        var slidePart = CopyTestSlidePart(reopenedDocument, targetSlide.Number);
        var audioId = copied.SdkOpenXmlElement.Descendants<A.AudioFromFile>().Single().Link!.Value;
        var mediaId = copied.SdkOpenXmlElement
            .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.Media>().Single().Embed!.Value;
        var audioRelationship = slidePart.DataPartReferenceRelationships.Single(r => r.Id == audioId);
        var mediaRelationship = slidePart.DataPartReferenceRelationships.Single(r => r.Id == mediaId);
        audioRelationship.Should().BeOfType<AudioReferenceRelationship>();
        audioRelationship.DataPart.Should().BeSameAs(mediaRelationship.DataPart);
        var imageId = copied.SdkOpenXmlElement.Descendants<A.Blip>().Single().Embed!.Value!;
        CopyTestPartBytes(slidePart.GetPartById(imageId)).Should().Equal(expectedImage);
        ValidatePresentation(reopened);
    }

    [TestCase("Object 2", 0)]
    [TestCase("Object 2", 1)]
    [TestCase("Object 2", 2)]
    [TestCase("Object 3", 0)]
    [TestCase("Object 3", 1)]
    [TestCase("Object 3", 2)]
    public void Add_copies_OLE_object_with_its_embedded_data(string shapeName, int destination)
    {
        // Arrange
        var source = new Presentation(TestAsset("009_table.pptx"));
        var original = source.Slide(2).Shapes.Shape(shapeName);
        using var sourceDocument = source.GetSdkPresentationDocument();
        var sourcePart = CopyTestSlidePart(sourceDocument, 2);
        var originalOle = original.SdkOpenXmlElement.Descendants<P.OleObject>().First();
        using var originalStream = sourcePart.GetPartById(originalOle.Id!.Value!).GetStream();
        using var expectedData = new MemoryStream();
        originalStream.CopyTo(expectedData);
        var originalImageId = original.SdkOpenXmlElement.Descendants<A.Blip>().First().Embed!.Value!;
        var expectedImage = CopyTestPartBytes(sourcePart.GetPartById(originalImageId));
        var target = destination == 2 ? new Presentation(p => p.Slide()) : source;
        var targetSlide = target.Slides[destination == 0 ? 1 : 0];
        var shapeCount = targetSlide.Shapes.Count;

        // Act
        targetSlide.Shapes.Add(original);

        // Assert
        using var reopened = SaveAndOpenPresentation(target);
        var copiedShapes = reopened.Slide(targetSlide.Number).Shapes;
        var copied = copiedShapes.Last();
        copiedShapes.Should().HaveCount(shapeCount + 1);
        copied.X.Should().Be(original.X);
        copied.Y.Should().Be(original.Y);
        var copiedOle = copied.SdkOpenXmlElement.Descendants<P.OleObject>().First();
        copiedOle.ProgId!.Value.Should().Be(originalOle.ProgId!.Value);
        using var reopenedDocument = reopened.GetSdkPresentationDocument();
        var targetPart = CopyTestSlidePart(reopenedDocument, targetSlide.Number);
        using var copiedStream = targetPart.GetPartById(copiedOle.Id!.Value!).GetStream();
        using var copiedData = new MemoryStream();
        copiedStream.CopyTo(copiedData);
        copiedData.ToArray().Should().Equal(expectedData.ToArray());
        var imageId = copied.SdkOpenXmlElement.Descendants<A.Blip>().First().Embed!.Value!;
        CopyTestPartBytes(targetPart.GetPartById(imageId)).Should().Equal(expectedImage);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Add_copies_video_to_another_presentation()
    {
        // Arrange
        var source = new Presentation(TestAsset("video-case001.pptx"));
        var original = source.Slide(1).Shape("Video 1");
        var expectedBytes = original.Media!.AsByteArray();
        var target = new Presentation(p => p.Slide());

        // Act
        target.Slide(1).Shapes.Add(original);

        // Assert
        using var reopened = SaveAndOpenPresentation(target);
        var copied = reopened.Slide(1).Shapes.Should().ContainSingle().Subject;
        copied.ContentType.Should().Be(ShapeContentType.Video);
        copied.Media!.MIME.Should().Be("video/mp4");
        copied.Media.AsByteArray().Should().Equal(expectedBytes);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Add_copies_audio_without_replacing_existing_media()
    {
        // Arrange
        var source = new Presentation(p => p.Slide());
        using var mp3 = TestAsset("064 mp3.mp3");
        source.Slide(1).Shapes.AddAudio(125, 75, mp3);
        var original = source.Slide(1).Shapes.Last();
        var target = new Presentation(p => p.Slide());
        using var wav = TestAsset("071 wav.wav");
        target.Slide(1).Shapes.AddAudio(25, 50, wav, AudioType.WAVE);
        var expectedExistingBytes = target.Slide(1).Shapes.Last().Media!.AsByteArray();

        // Act
        target.Slide(1).Shapes.Add(original);

        // Assert
        using var reopened = SaveAndOpenPresentation(target);
        var shapes = reopened.Slide(1).Shapes;
        shapes.Should().HaveCount(2);
        shapes[0].Media!.MIME.Should().Be("audio/wav");
        shapes[0].Media.AsByteArray().Should().Equal(expectedExistingBytes);
        shapes[1].Media!.MIME.Should().Be("audio/mpeg");
        shapes[1].Media.AsByteArray().Should().Equal(original.Media!.AsByteArray());
        ValidatePresentation(reopened);
    }

    private static SlidePart CopyTestSlidePart(PresentationDocument document, int slideNumber)
    {
        var presentationPart = document.PresentationPart!;
        var slideId = presentationPart.Presentation.SlideIdList!.Elements<P.SlideId>().ElementAt(slideNumber - 1);
        return (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!.Value!);
    }

    private static byte[] CopyTestPartBytes(OpenXmlPart part)
    {
        using var stream = part.GetStream();
        using var bytes = new MemoryStream();
        stream.CopyTo(bytes);
        return bytes.ToArray();
    }
}
