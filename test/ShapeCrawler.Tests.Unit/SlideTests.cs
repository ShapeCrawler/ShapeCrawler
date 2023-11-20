using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class SlideTests : SCTest
{
    [Test]
    public void Hide_MethodHidesSlide_WhenItIsExecuted()
    {
        // Arrange
        var pptx = StreamOf("001.pptx");
        var pre = new Presentation(pptx);
        var slide = pre.Slides.First();

        // Act
        slide.Hide();

        // Assert
        slide.Hidden().Should().Be(true);
    }

    [Test]
    public void Hidden_GetterReturnsTrue_WhenTheSlideIsHidden()
    {
        // Arrange
        var pptx = StreamOf("002.pptx");
        var pres = new Presentation(pptx);
        ISlide slideEx = pres.Slides[2];

        // Act
        bool hidden = slideEx.Hidden();

        // Assert
        hidden.Should().BeTrue();
    }

    [Test]
    public void Background_SetImage_updates_background()
    {
        // Arrange
        var pre = new Presentation(StreamOf("009_table.pptx"));
        var backgroundImage = pre.Slides[0].Background;
        var image = StreamOf("test-image-2.png");
        var bytesBefore = backgroundImage.AsByteArray();

        // Act
        backgroundImage.Update(image);

        // Assert
        var bytesAfter = backgroundImage.AsByteArray();
        bytesAfter.Length.Should().NotBe(bytesBefore.Length);
    }

    [Test]
    public void Background_SetImage_updates_background_of_new_slide()
    {
        // Arrange
        var pres = new Presentation();
        pres.Slides.AddEmptySlide(SlideLayoutType.Blank);
        var slide = pres.Slides[0];
        var bgImage = StreamOf("test-image-2.png");

        // Act
        slide.Background.Update(bgImage);

        // Assert
        slide.Background.Should().NotBeNull();
    }

    [Test]
    public void Background_AsByteArray_throws_exception_slide_doesnt_have_background()
    {
        // Arrange
        var pres = new Presentation(StreamOf("009_table.pptx"));
        var slide = pres.Slides[1];

        // Act
        var act = () => slide.Background.AsByteArray();

        // Assert
        act.Should().Throw<Exception>();
    }

    [Test]
    public void CustomData_ReturnsData_WhenCustomDataWasAssigned()
    {
        // Arrange
        const string customDataString = "Test custom data";
        var originPre = new Presentation(StreamOf("001.pptx"));
        var slide = originPre.Slides.First();

        // Act
        slide.CustomData = customDataString;

        var savedPreStream = new MemoryStream();
        originPre.SaveAs(savedPreStream);
        var savedPre = new Presentation(savedPreStream);
        var customData = savedPre.Slides.First().CustomData;

        // Assert
        customData.Should().Be(customDataString);
    }

    [Test]
    public void CustomData_PropertyIsNull_WhenTheSlideHasNotCustomData()
    {
        // Arrange
        var slide = new Presentation(StreamOf("001.pptx")).Slides.First();

        // Act
        var sldCustomData = slide.CustomData;

        // Assert
        sldCustomData.Should().BeNull();
    }

    [Test]
    public void Number_Setter_moves_slide_to_specified_number_position()
    {
        // Arrange
        var pptxStream = StreamOf("001.pptx");
        var pres = new Presentation(pptxStream);
        var slide1 = pres.Slides[0];
        var slide2 = pres.Slides[1];
        slide1.CustomData = "old-number-1";

        // Act
        slide1.Number = 2;

        // Assert
        slide1.Number.Should().Be(2);
        slide2.Number.Should().Be(1, "because the first slide was inserted to its position.");

        pres.Save();
        pres = new Presentation(pptxStream);
        slide2 = pres.Slides.First(s => s.CustomData == "old-number-1");
        slide2.Number.Should().Be(2);
    }

    [Test]
    public void Number_Setter()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];

        // Act
        slide.Number = 1;

        // Assert
        slide.Number.Should().Be(1);
    }

    [Test]
    public void GetAllTextboxes_contains_all_textboxes_withTable()
    {
        // Arrange
        var pptx = StreamOf("039.pptx");
        var pres = new Presentation(pptx);
        var slide = pres.Slides.First();

        // Act
        var textboxes = slide.TextFrames();

        // Assert
        textboxes.Count.Should().Be(11);
    }

    [Test]
    public void TextFrames_returns_list_of_all_text_frames_on_that_slide()
    {
        // Arrange
        var pres = new Presentation(StreamOf("011_dt.pptx"));
        var slide = pres.Slides.First();

        // Act
        var textFrames = slide.TextFrames();

        // Assert
        textFrames.Count.Should().Be(4);
    }

#if DEBUG
    [Fact(Skip = "In progress")]
    public void SaveAsPng_saves_slide_as_image()
    {
        // Arrange
        var pptxStream = StreamOf("autoshape-case011_save-as-png.pptx");
        var pres = new Presentation(pptxStream);
        var slide = pres.Slides[0];
        var mStream = new MemoryStream();

        // Act
        slide.SaveAsPng(mStream);
    }

    [Test]
    public void ToHTML_converts_slide_to_HTML()
    {
        // Arrange
        var pptx = TestHelper.GetStream("autoshape-case011_save-as-png.pptx");
        var pre = new Presentation(pptx);
        var slide = pre.Slides[0];

        // Act
        var slideHtml = slide.ToHtml();

        // Arrange
    }

#endif
}