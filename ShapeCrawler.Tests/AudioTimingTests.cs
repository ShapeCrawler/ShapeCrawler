using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;

namespace ShapeCrawler.Tests;

[TestFixture]
public class AudioTimingTests : SCTest
{
    [Test]
    public void StartDelay_defaults_to_zero()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        presentation.Slide(1).Shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));

        // Act
        var delay = presentation.Slide(1).Shapes.Last().Media!.StartDelay;

        // Assert
        delay.Should().Be(TimeSpan.Zero);
    }

    [TestCase(AudioStartMode.InClickSequence)]
    [TestCase(AudioStartMode.Automatically)]
    [TestCase(AudioStartMode.WhenClickedOn)]
    public void StartDelay_sets_effect_delay_and_preserves_start_mode(AudioStartMode mode)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;
        shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));
        var media = shapes.Last().Media!;
        media.StartMode = mode;

        // Act
        media.StartDelay = TimeSpan.FromSeconds(2);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        var savedMedia = reopened.Slide(1).Shapes.Last().Media!;
        savedMedia.StartDelay.Should().Be(TimeSpan.FromSeconds(2));
        savedMedia.StartMode.Should().Be(mode);
        ValidatePresentation(reopened);
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var effect = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<Command>().Single()
            .Ancestors<CommonTimeNode>().First(node => node.PresetClass?.Value == TimeNodePresetClassValues.MediaCall);
        effect.StartConditionList!.GetFirstChild<Condition>()!.Delay!.Value.Should().Be("2000");
        effect.NodeType!.Value.Should().Be(mode == AudioStartMode.InClickSequence
            ? TimeNodeValues.ClickEffect : TimeNodeValues.WithEffect);
    }

    [TestCase(AudioStartMode.InClickSequence)]
    [TestCase(AudioStartMode.Automatically)]
    [TestCase(AudioStartMode.WhenClickedOn)]
    public void StartMode_preserves_start_delay(AudioStartMode mode)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        presentation.Slide(1).Shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));
        var media = presentation.Slide(1).Shapes.Last().Media!;
        media.StartDelay = TimeSpan.FromSeconds(2);

        // Act
        media.StartMode = AudioStartMode.WhenClickedOn;
        media.StartMode = AudioStartMode.Automatically;
        media.StartMode = mode;

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.Last().Media!.StartDelay.Should().Be(TimeSpan.FromSeconds(2));
        ValidatePresentation(reopened);
    }

    [TestCase(0d, "0")]
    [TestCase(1234.6d, "1235")]
    [TestCase(4294967295d, "4294967295")]
    public void StartDelay_saves_millisecond_precision(double milliseconds, string expectedDelay)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        presentation.Slide(1).Shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));
        var media = presentation.Slide(1).Shapes.Last().Media!;
        media.StartDelay = TimeSpan.FromSeconds(2);

        // Act
        media.StartDelay = TimeSpan.FromMilliseconds(milliseconds);

        // Assert
        using var sdk = SaveAndOpenPresentationAsSdk(presentation);
        var effect = sdk.PresentationPart!.SlideParts.Single().Slide.Descendants<Command>().Single()
            .Ancestors<CommonTimeNode>().First(node => node.PresetClass?.Value == TimeNodePresetClassValues.MediaCall);
        effect.StartConditionList!.GetFirstChild<Condition>()!.Delay!.Value.Should().Be(expectedDelay);
        ValidatePresentation(presentation);
    }

    [TestCase(-1d)]
    [TestCase(4294967296d)]
    public void StartDelay_rejects_invalid_value_without_changing_timing(double milliseconds)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        presentation.Slide(1).Shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));
        var media = presentation.Slide(1).Shapes.Last().Media!;
        using var original = SaveAndOpenPresentationAsSdk(presentation);
        var originalTiming = original.PresentationPart!.SlideParts.Single().Slide.Timing!.OuterXml;

        // Act
        var act = () => media.StartDelay = TimeSpan.FromMilliseconds(milliseconds);

        // Assert
        act.Should().Throw<ArgumentOutOfRangeException>();
        using var updated = SaveAndOpenPresentationAsSdk(presentation);
        updated.PresentationPart!.SlideParts.Single().Slide.Timing!.OuterXml.Should().Be(originalTiming);
    }

    [Test]
    public void StartDelay_preserves_existing_PowerPoint_timing_and_other_audio()
    {
        // Arrange
        using var presentation = new Presentation(TestAsset("audio-case001.pptx"));
        var shapes = presentation.Slide(1).Shapes;
        var media = shapes.First(shape => shape.Media != null).Media!;
        shapes.AddAudio(300, 100, TestAsset("071 wav.wav"), AudioType.WAVE);
        shapes.Last().Media!.StartDelay = TimeSpan.FromSeconds(3);
        using var original = SaveAndOpenPresentationAsSdk(presentation);
        var expected = original.PresentationPart!.SlideParts.First().Slide.Timing!;
        var effect = expected.Descendants<Command>().First().Ancestors<CommonTimeNode>()
            .First(node => node.PresetClass?.Value == TimeNodePresetClassValues.MediaCall);
        effect.StartConditionList!.GetFirstChild<Condition>()!.Delay = "2000";

        // Act
        media.StartDelay = TimeSpan.FromSeconds(2);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.First(shape => shape.Media != null).Media!.StartDelay.Should().Be(TimeSpan.FromSeconds(2));
        reopened.Slide(1).Shapes.Last().Media!.StartDelay.Should().Be(TimeSpan.FromSeconds(3));
        using var updated = SaveAndOpenPresentationAsSdk(reopened);
        updated.PresentationPart!.SlideParts.First().Slide.Timing!.OuterXml.Should().Be(expected.OuterXml);
        ValidatePresentation(reopened);
    }

    [Test]
    public void StartDelay_initializes_missing_timing()
    {
        // Arrange
        using var source = new Presentation(p => p.Slide());
        source.Slide(1).Shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));
        using var sdk = SaveAndOpenPresentationAsSdk(source);
        sdk.PresentationPart!.SlideParts.Single().Slide.Timing!.Remove();
        using var stream = new MemoryStream();
        using (var clone = sdk.Clone(stream)) { }
        stream.Position = 0;
        using var presentation = new Presentation(stream);
        var media = presentation.Slide(1).Shapes.Last().Media!;

        // Act
        media.StartDelay = TimeSpan.FromSeconds(2);

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.Last().Media!.StartDelay.Should().Be(TimeSpan.FromSeconds(2));
        reopened.Slide(1).Shapes.Last().Media!.StartMode.Should().Be(AudioStartMode.WhenClickedOn);
        ValidatePresentation(reopened);
    }
}
