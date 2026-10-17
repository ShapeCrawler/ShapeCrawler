using DocumentFormat.OpenXml.Presentation;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;

namespace ShapeCrawler.Tests;

[TestFixture]
public class AudioStartModeTests : SCTest
{
    [Test]
    public void AddAudio_defaults_to_InClickSequence()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());

        // Act
        presentation.Slide(1).Shapes.AddAudio(300, 100, TestAsset("064 mp3.mp3"));

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);

        reopened.Slide(1).Shapes.Last().Media!.StartMode.Should().Be(AudioStartMode.InClickSequence);
        ValidatePresentation(reopened);
    }

    [TestCase(AudioStartMode.InClickSequence)]
    [TestCase(AudioStartMode.Automatically)]
    [TestCase(AudioStartMode.WhenClickedOn)]
    public void StartMode_sets_single_audio_playback_trigger(AudioStartMode mode)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        presentation.Slide(1).Shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));

        // Act
        presentation.Slide(1).Shapes.Last().Media!.StartMode = mode;

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.Last().Media!.StartMode.Should().Be(mode);
        ValidatePresentation(reopened);
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var timing = sdk.PresentationPart!.SlideParts.Single().Slide.Timing!;
        var command = timing.Descendants<Command>().Single();
        var effect = command.Ancestors<CommonTimeNode>()
            .First(node => node.PresetClass?.Value == TimeNodePresetClassValues.MediaCall);
        effect.NodeType!.Value.Should().Be(mode == AudioStartMode.InClickSequence
            ? TimeNodeValues.ClickEffect : TimeNodeValues.WithEffect);
        var sequence = command.Ancestors<CommonTimeNode>()
            .First(node => node.NodeType?.Value == TimeNodeValues.MainSequence ||
                node.NodeType?.Value == TimeNodeValues.InteractiveSequence);
        if (mode == AudioStartMode.WhenClickedOn)
        {
            sequence.StartConditionList!.Descendants<ShapeTarget>().Single().ShapeId!.Value
                .Should().Be(command.Descendants<ShapeTarget>().Single().ShapeId!.Value);
            sequence.StartConditionList.GetFirstChild<Condition>()!.Event!.Value.Should().Be(TriggerEventValues.OnClick);
        }
    }

    [TestCase(AudioStartMode.InClickSequence)]
    [TestCase(AudioStartMode.Automatically)]
    [TestCase(AudioStartMode.WhenClickedOn)]
    public void StartMode_round_trips_and_preserves_other_audio(AudioStartMode mode)
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        var shapes = presentation.Slide(1).Shapes;
        shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));
        shapes.AddAudio(300, 100, TestAsset("071 wav.wav"), AudioType.WAVE);
        var media = shapes.Last().Media!;

        // Act
        media.StartMode = AudioStartMode.WhenClickedOn;
        media.StartMode = AudioStartMode.Automatically;
        media.StartMode = mode;

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);

        reopened.Slide(1).Shapes.Last().Media!.StartMode.Should().Be(mode);
        reopened.Slide(1).Shapes.First(s => s.Media != null).Media!.StartMode.Should().Be(AudioStartMode.InClickSequence);
        ValidatePresentation(reopened);
        using var sdk = SaveAndOpenPresentationAsSdk(reopened);
        var timing = sdk.PresentationPart!.SlideParts.Single().Slide.Timing!;
        var ids = timing.Descendants<CommonTimeNode>().Select(node => node.Id!.Value).ToArray();
        ids.Should().OnlyHaveUniqueItems();
        timing.Descendants<Audio>().Should().HaveCount(2);
        timing.Descendants<Command>().Should().HaveCount(2);
    }

    [Test]
    public void StartMode_reads_and_updates_existing_PowerPoint_timing()
    {
        // Arrange
        using var presentation = new Presentation(TestAsset("audio-case001.pptx"));
        var media = presentation.Slide(1).Shapes.First(s => s.Media != null).Media!;
        media.StartMode.Should().Be(AudioStartMode.InClickSequence);
        using var originalSdk = SaveAndOpenPresentationAsSdk(presentation);
        var originalCommand = originalSdk.PresentationPart!.SlideParts.First().Slide
            .Descendants<Command>().Single();
        var originalDuration = originalCommand.Descendants<CommonTimeNode>().Single().Duration!.Value;

        // Act
        media.StartMode = AudioStartMode.Automatically;

        // Assert
        using var reopened = SaveAndOpenPresentation(presentation);
        reopened.Slide(1).Shapes.First(s => s.Media != null).Media!.StartMode.Should().Be(AudioStartMode.Automatically);
        ValidatePresentation(reopened);
        using var updatedSdk = SaveAndOpenPresentationAsSdk(reopened);
        var updatedCommand = updatedSdk.PresentationPart!.SlideParts.First().Slide.Descendants<Command>().Single();
        updatedCommand.Descendants<CommonTimeNode>().Single().Duration!.Value.Should().Be(originalDuration);
    }

    [Test]
    public void StartMode_rejects_invalid_value_without_changing_timing()
    {
        // Arrange
        using var presentation = new Presentation(p => p.Slide());
        presentation.Slide(1).Shapes.AddAudio(100, 100, TestAsset("064 mp3.mp3"));
        var media = presentation.Slide(1).Shapes.Last().Media!;

        // Act
        var act = () => media.StartMode = (AudioStartMode)99;

        // Assert
        act.Should().Throw<ArgumentOutOfRangeException>();
        media.StartMode.Should().Be(AudioStartMode.InClickSequence);
    }
}
