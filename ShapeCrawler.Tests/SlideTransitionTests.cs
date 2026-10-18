using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests;

public class SlideTransitionTests : SCTest
{
    [TestCase(false)]
    [TestCase(true)]
    public void Getters_return_defaults_without_changing_slide_xml(bool hasTransition)
    {
        // Arrange
        using var source = SlideXml(slide =>
        {
            if (hasTransition)
            {
                slide.AddChild(new P.Transition(), true);
            }
        });
        using var pres = new Presentation(source);
        string before;
        using (var document = PresentationDocument.Open(source, false))
        {
            before = document.PresentationPart!.SlideParts.Single().Slide.OuterXml;
        }

        // Act
        var transition = pres.Slides[0].Transition;
        var click = transition.AdvanceOnClick;
        var time = transition.AdvanceAfterTime;

        // Assert
        click.Should().BeTrue();
        time.Should().BeNull();
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var result = PresentationDocument.Open(saved, false);
        result.PresentationPart!.SlideParts.Single().Slide.OuterXml.Should().Be(before);
    }

    [TestCase(false, null)]
    [TestCase(true, null)]
    [TestCase(false, 0U)]
    [TestCase(true, 0U)]
    [TestCase(false, 5000U)]
    [TestCase(true, 5000U)]
    [TestCase(false, 2147483647U)]
    public void Advance_settings_round_trip_independently(bool click, uint? milliseconds)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var transition = pres.Slides[0].Transition;
        transition.AdvanceOnClick = !click;
        transition.AdvanceAfterTime = 123;

        // Act
        transition.AdvanceOnClick = click;
        transition.AdvanceAfterTime = milliseconds;

        // Assert
        transition.AdvanceOnClick.Should().Be(click);
        transition.AdvanceAfterTime.Should().Be(milliseconds);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var xml = document.PresentationPart!.SlideParts.Single().Slide.Transition!;
            xml.AdvanceOnClick!.Value.Should().Be(click);
            xml.AdvanceAfterTime?.Value.Should().Be(milliseconds?.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        using var reopened = new Presentation(saved);
        reopened.Slides[0].Transition.AdvanceOnClick.Should().Be(click);
        reopened.Slides[0].Transition.AdvanceAfterTime.Should().Be(milliseconds);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Clearing_absent_timer_does_not_create_transition()
    {
        // Arrange
        using var source = SlideXml(_ => { });
        using var pres = new Presentation(source);

        // Act
        pres.Slides[0].Transition.AdvanceAfterTime = null;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var document = PresentationDocument.Open(saved, false);
        document.PresentationPart!.SlideParts.Single().Slide.Transition.Should().BeNull();
    }

    [Test]
    public void Existing_effect_speed_and_animation_are_preserved_when_clearing_timer()
    {
        // Arrange
        using var source = SlideXml(slide =>
        {
            slide.AddChild(new P.Transition(new P.FadeTransition())
            {
                Speed = P.TransitionSpeedValues.Slow,
                AdvanceAfterTime = "1000"
            }, true);
            slide.AddChild(new P.Timing("""
                <p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
                  <p:tnLst><p:par><p:cTn id="1" dur="indefinite" nodeType="tmRoot"/></p:par></p:tnLst>
                </p:timing>
                """), true);
        });
        using var pres = new Presentation(source);
        string timingBefore;
        using (var document = PresentationDocument.Open(source, false))
        {
            timingBefore = document.PresentationPart!.SlideParts.Single().Slide.Timing!.OuterXml;
        }

        // Act
        pres.Slides[0].Transition.AdvanceOnClick = false;
        pres.Slides[0].Transition.AdvanceAfterTime = null;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide;
            slide.Transition!.GetFirstChild<P.FadeTransition>().Should().NotBeNull();
            slide.Transition.Speed!.Value.Should().Be(P.TransitionSpeedValues.Slow);
            slide.Transition.AdvanceAfterTime.Should().BeNull();
            slide.Timing!.OuterXml.Should().Be(timingBefore);
        }

        ValidatePresentation(pres);
    }

    [Test]
    public void New_transition_is_inserted_before_timing_and_extensions()
    {
        // Arrange
        using var source = SlideXml(slide =>
        {
            slide.AddChild(new P.Timing(), true);
            slide.AddChild(new P.SlideExtensionList(), true);
        });
        using var pres = new Presentation(source);

        // Act
        pres.Slides[0].Transition.AdvanceAfterTime = 2500;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide;
            var names = slide.ChildElements.Select(element => element.LocalName).ToList();
            names.IndexOf("transition").Should().BeLessThan(names.IndexOf("timing"));
            names.IndexOf("timing").Should().BeLessThan(names.IndexOf("extLst"));
        }

        ValidatePresentation(pres);
    }

    [TestCase(null)]
    [TestCase(4000U)]
    public void Setters_update_choice_and_fallback_without_adding_duplicate_transition(uint? milliseconds)
    {
        // Arrange
        const string xml = """
            <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
              <mc:Choice Requires="p14"><p:transition advTm="1000" p14:dur="700"><p:fade/></p:transition></mc:Choice>
              <mc:Fallback><p:transition advTm="1000"><p:fade/></p:transition></mc:Fallback>
            </mc:AlternateContent>
            """;
        using var source = SlideXml(slide => slide.AppendChild(new AlternateContent(xml)));
        using var pres = new Presentation(source);

        // Act
        pres.Slides[0].Transition.AdvanceOnClick = false;
        pres.Slides[0].Transition.AdvanceAfterTime = milliseconds;
        pres.Slides[0].Transition.Type = SlideTransitionType.Push;
        pres.Slides[0].Transition.Speed = SlideTransitionSpeed.Medium;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide;
            slide.Elements<P.Transition>().Should().BeEmpty();
            var transitions = slide.Descendants<P.Transition>().ToArray();
            transitions.Should().HaveCount(2);
            foreach (var transition in transitions)
            {
                transition.AdvanceOnClick!.Value.Should().BeFalse();
                transition.AdvanceAfterTime?.Value.Should().Be(milliseconds?.ToString(System.Globalization.CultureInfo.InvariantCulture));
                transition.GetFirstChild<P.PushTransition>().Should().NotBeNull();
                transition.Speed!.Value.Should().Be(P.TransitionSpeedValues.Medium);
                transition.GetAttributes().Should().NotContain(attribute => attribute.LocalName == "dur");
            }
        }

        using var reopened = new Presentation(saved);
        reopened.Slides[0].Transition.AdvanceOnClick.Should().BeFalse();
        reopened.Slides[0].Transition.AdvanceAfterTime.Should().Be(milliseconds);
        reopened.Slides[0].Transition.Type.Should().Be(SlideTransitionType.Push);
        reopened.Slides[0].Transition.Speed.Should().Be(SlideTransitionSpeed.Medium);
        ValidatePresentation(reopened);
    }

    [TestCase(2147483648U)]
    [TestCase(uint.MaxValue)]
    public void AdvanceAfterTime_rejects_out_of_range_values_without_changing_settings(uint milliseconds)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var transition = pres.Slides[0].Transition;
        transition.AdvanceAfterTime = 5000;

        // Act
        var setTime = () => transition.AdvanceAfterTime = milliseconds;

        // Assert
        setTime.Should().Throw<ArgumentOutOfRangeException>();
        transition.AdvanceAfterTime.Should().Be(5000);
        transition.AdvanceOnClick.Should().BeTrue();
        ValidatePresentation(pres);
    }

    [TestCase(SlideTransitionType.None, null)]
    [TestCase(SlideTransitionType.Blinds, "blinds")]
    [TestCase(SlideTransitionType.Checker, "checker")]
    [TestCase(SlideTransitionType.Circle, "circle")]
    [TestCase(SlideTransitionType.Comb, "comb")]
    [TestCase(SlideTransitionType.Cover, "cover")]
    [TestCase(SlideTransitionType.Cut, "cut")]
    [TestCase(SlideTransitionType.Diamond, "diamond")]
    [TestCase(SlideTransitionType.Dissolve, "dissolve")]
    [TestCase(SlideTransitionType.Fade, "fade")]
    [TestCase(SlideTransitionType.Newsflash, "newsflash")]
    [TestCase(SlideTransitionType.Plus, "plus")]
    [TestCase(SlideTransitionType.Pull, "pull")]
    [TestCase(SlideTransitionType.Push, "push")]
    [TestCase(SlideTransitionType.Random, "random")]
    [TestCase(SlideTransitionType.RandomBar, "randomBar")]
    [TestCase(SlideTransitionType.Split, "split")]
    [TestCase(SlideTransitionType.Strips, "strips")]
    [TestCase(SlideTransitionType.Wedge, "wedge")]
    [TestCase(SlideTransitionType.Wheel, "wheel")]
    [TestCase(SlideTransitionType.Wipe, "wipe")]
    [TestCase(SlideTransitionType.Zoom, "zoom")]
    public void Type_round_trips_standard_effects_without_changing_advance_settings(SlideTransitionType type, string? elementName)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        var transition = pres.Slides[0].Transition;
        transition.Type = SlideTransitionType.Fade;
        transition.AdvanceOnClick = false;
        transition.AdvanceAfterTime = 5000;

        // Act
        transition.Type = type;

        // Assert
        transition.Type.Should().Be(type);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var effects = document.PresentationPart!.SlideParts.Single().Slide.Transition!.ChildElements;
            effects.Select(effect => effect.LocalName).Should().Equal(elementName is null ? [] : new[] { elementName });
        }

        using var reopened = new Presentation(saved);
        reopened.Slides[0].Transition.Type.Should().Be(type);
        reopened.Slides[0].Transition.AdvanceOnClick.Should().BeFalse();
        reopened.Slides[0].Transition.AdvanceAfterTime.Should().Be(5000);
        ValidatePresentation(reopened);
    }

    [TestCase(SlideTransitionSpeed.Fast, "fast")]
    [TestCase(SlideTransitionSpeed.Medium, "med")]
    [TestCase(SlideTransitionSpeed.Slow, "slow")]
    public void Speed_round_trips_and_clears_explicit_duration(SlideTransitionSpeed speed, string xmlValue)
    {
        // Arrange
        using var source = SlideXml(slide =>
        {
            var xml = new P.Transition(new P.FadeTransition()) { AdvanceAfterTime = "5000" };
            xml.SetAttribute(new OpenXmlAttribute("p14", "dur", "http://schemas.microsoft.com/office/powerpoint/2010/main", "700"));
            slide.AddChild(xml, true);
        });
        using var pres = new Presentation(source);

        // Act
        pres.Slides[0].Transition.Speed = speed;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var transition = document.PresentationPart!.SlideParts.Single().Slide.Transition!;
            transition.GetAttribute("spd", string.Empty).Value.Should().Be(xmlValue);
            transition.GetAttributes().Should().NotContain(attribute => attribute.LocalName == "dur");
            transition.GetFirstChild<P.FadeTransition>().Should().NotBeNull();
        }

        using var reopened = new Presentation(saved);
        reopened.Slides[0].Transition.Speed.Should().Be(speed);
        reopened.Slides[0].Transition.AdvanceAfterTime.Should().Be(5000);
        ValidatePresentation(reopened);
    }

    [Test]
    public void Type_and_speed_defaults_do_not_create_xml()
    {
        // Arrange
        using var source = SlideXml(_ => { });
        using var pres = new Presentation(source);

        // Act
        var type = pres.Slides[0].Transition.Type;
        var speed = pres.Slides[0].Transition.Speed;

        // Assert
        type.Should().Be(SlideTransitionType.None);
        speed.Should().Be(SlideTransitionSpeed.Fast);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var document = PresentationDocument.Open(saved, false);
        document.PresentationPart!.SlideParts.Single().Slide.Transition.Should().BeNull();
    }

    [TestCase(SlideTransitionType.Wheel)]
    [TestCase(SlideTransitionType.None)]
    [TestCase(SlideTransitionType.Push)]
    public void Type_preserves_sound_extensions_and_animation(SlideTransitionType type)
    {
        // Arrange
        using var source = SlideXml(slide =>
        {
            var transition = new P.Transition(new P.WheelTransition { Spokes = 8 },
                new P.SoundAction(new P.EndSoundAction()), new P.ExtensionListWithModification())
            {
                Speed = P.TransitionSpeedValues.Slow,
                AdvanceAfterTime = "5000"
            };
            transition.SetAttribute(new OpenXmlAttribute("p14", "dur", "http://schemas.microsoft.com/office/powerpoint/2010/main", "700"));
            slide.AddChild(transition, true);
            slide.AddChild(new P.Timing(), true);
        });
        using var pres = new Presentation(source);

        // Act
        pres.Slides[0].Transition.Type = type;

        // Assert
        using var saved = new MemoryStream();
        pres.Save(saved);
        using (var document = PresentationDocument.Open(saved, false))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide;
            var transition = slide.Transition!;
            transition.GetFirstChild<P.SoundAction>()!.GetFirstChild<P.EndSoundAction>().Should().NotBeNull();
            transition.GetFirstChild<P.ExtensionListWithModification>().Should().NotBeNull();
            slide.Timing.Should().NotBeNull();
            if (type == SlideTransitionType.Wheel)
            {
                transition.GetFirstChild<P.WheelTransition>()!.Spokes!.Value.Should().Be(8);
                transition.GetAttributes().Should().Contain(attribute => attribute.LocalName == "dur" && attribute.Value == "700");
            }
            else
            {
                transition.GetAttributes().Should().NotContain(attribute => attribute.LocalName == "dur");
            }
        }

        pres.Slides[0].Transition.Speed.Should().Be(SlideTransitionSpeed.Slow);
        pres.Slides[0].Transition.AdvanceAfterTime.Should().Be(5000);
        ValidatePresentation(pres);
    }

    [TestCase(SlideTransitionType.Unsupported)]
    [TestCase((SlideTransitionType)999)]
    public void Type_rejects_invalid_values_before_changing_document(SlideTransitionType type)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Transition.Type = SlideTransitionType.Push;

        // Act
        var setType = () => pres.Slides[0].Transition.Type = type;

        // Assert
        setType.Should().Throw<ArgumentOutOfRangeException>();
        pres.Slides[0].Transition.Type.Should().Be(SlideTransitionType.Push);
    }

    [Test]
    public void Speed_rejects_invalid_values_before_changing_document()
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Transition.Speed = SlideTransitionSpeed.Slow;

        // Act
        var setSpeed = () => pres.Slides[0].Transition.Speed = (SlideTransitionSpeed)999;

        // Assert
        setSpeed.Should().Throw<ArgumentOutOfRangeException>();
        pres.Slides[0].Transition.Speed.Should().Be(SlideTransitionSpeed.Slow);
    }

    [Test]
    public void Newer_effect_is_reported_as_unsupported_and_preserved_when_setting_speed()
    {
        // Arrange
        using var source = SlideXml(slide => slide.AddChild(new P.Transition("""
            <p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:flash/></p:transition>
            """), true));
        using var pres = new Presentation(source);

        // Act
        var type = pres.Slides[0].Transition.Type;
        pres.Slides[0].Transition.Speed = SlideTransitionSpeed.Slow;

        // Assert
        type.Should().Be(SlideTransitionType.Unsupported);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var document = PresentationDocument.Open(saved, false);
        document.PresentationPart!.SlideParts.Single().Slide.Transition!.FirstChild!.LocalName.Should().Be("flash");
        ValidatePresentation(pres);
    }

    private static MemoryStream SlideXml(Action<P.Slide> edit)
    {
        using var pres = new Presentation(p => p.Slide());
        var stream = new MemoryStream();
        pres.Save(stream);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide;
            slide.Transition?.Remove();
            edit(slide);
        }

        stream.Position = 0;
        return stream;
    }
}
