using FluentAssertions;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

[TestFixture]
public class MediaTests : SCTest
{
    [Test]
    public void Update_replaces_audio_shape_content()
    {
        // Arrange
        var originalAudio = TestAsset("064 mp3.mp3");
        var replacingAudio = TestAsset("071 wav.wav");
        var replacingAudioBytes = replacingAudio.ToArray();
        var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;
        shapes.AddAudio(x: 111, y: 222, originalAudio, AudioType.MP3);
        var audioShape = shapes.Last();
        audioShape.Width = 73;
        audioShape.Height = 41;

        // Act
        audioShape.Media!.Update(replacingAudio);
        presentation = SaveAndOpenPresentation(presentation);
        audioShape = presentation.Slide(1).Shapes.Last();

        // Assert
        audioShape.ContentType.Should().Be(ShapeContentType.Audio);
        audioShape.Media!.MIME.Should().Be("audio/wav");
        audioShape.Media.AsByteArray().Should().Equal(replacingAudioBytes);
        audioShape.X.Should().Be(111);
        audioShape.Y.Should().Be(222);
        audioShape.Width.Should().Be(73);
        audioShape.Height.Should().Be(41);
        ValidatePresentation(presentation);
    }
}