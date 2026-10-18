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
              <mc:Choice Requires="p14"><p:transition advTm="1000"><p:fade/></p:transition></mc:Choice>
              <mc:Fallback><p:transition advTm="1000"><p:fade/></p:transition></mc:Fallback>
            </mc:AlternateContent>
            """;
        using var source = SlideXml(slide => slide.AppendChild(new AlternateContent(xml)));
        using var pres = new Presentation(source);

        // Act
        pres.Slides[0].Transition.AdvanceOnClick = false;
        pres.Slides[0].Transition.AdvanceAfterTime = milliseconds;

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
                transition.GetFirstChild<P.FadeTransition>().Should().NotBeNull();
            }
        }

        using var reopened = new Presentation(saved);
        reopened.Slides[0].Transition.AdvanceOnClick.Should().BeFalse();
        reopened.Slides[0].Transition.AdvanceAfterTime.Should().Be(milliseconds);
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
