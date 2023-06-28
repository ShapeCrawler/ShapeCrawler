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
    [Fact]
    public void Hide_MethodHidesSlide_WhenItIsExecuted()
    {
        // Arrange
        var pptx = GetTestStream("001.pptx");
        var pre = SCPresentation.Open(pptx);
        var slide = pre.Slides.First();

        // Act
        slide.Hide();

        // Assert
        slide.Hidden.Should().Be(true);
    }

    [Fact]
    public void Hidden_GetterReturnsTrue_WhenTheSlideIsHidden()
    {
        // Arrange
        var pptx = GetTestStream("002.pptx");
        var pres = SCPresentation.Open(pptx);
        ISlide slideEx = pres.Slides[2];

        // Act
        bool hidden = slideEx.Hidden;

        // Assert
        hidden.Should().BeTrue();
    }

    [Fact]
    public async void Background_SetImage_updates_background()
    {
        // Arrange
        var pptx = GetTestStream("009_table.pptx");
var pre = SCPresentation.Open(pptx);
        var backgroundImage = pre.Slides[0].Background;
        var imgStream = new MemoryStream(Resources.test_image_2);
        var bytesBefore = await backgroundImage.BinaryData.ConfigureAwait(false);

        // Act
        backgroundImage.SetImage(imgStream);

        // Assert
        var bytesAfter = await backgroundImage.BinaryData.ConfigureAwait(false);
        bytesAfter.Length.Should().NotBe(bytesBefore.Length);
    }

    [Fact]
    public void Background_ImageIsNull_WhenTheSlideHasNotBackground()
    {
        // Arrange
        var slide = SCPresentation.Open(GetTestStream("009_table.pptx")).Slides[1];

        // Act
        var backgroundImage = slide.Background;

        // Assert
        backgroundImage.Should().BeNull();
    }

    [Fact]
    public void CustomData_ReturnsData_WhenCustomDataWasAssigned()
    {
        // Arrange
        const string customDataString = "Test custom data";
        var originPre = SCPresentation.Open(GetTestStream("001.pptx"));
        var slide = originPre.Slides.First();

        // Act
        slide.CustomData = customDataString;

        var savedPreStream = new MemoryStream();
        originPre.SaveAs(savedPreStream);
        var savedPre = SCPresentation.Open(savedPreStream);
        var customData = savedPre.Slides.First().CustomData;

        // Assert
        customData.Should().Be(customDataString);
    }

    [Fact]
    public void CustomData_PropertyIsNull_WhenTheSlideHasNotCustomData()
    {
        // Arrange
        var slide = SCPresentation.Open(GetTestStream("001.pptx")).Slides.First();

        // Act
        var sldCustomData = slide.CustomData;

        // Assert
        sldCustomData.Should().BeNull();
    }

    [Fact]
    public void Number_Setter_moves_slide_to_specified_number_position()
    {
        // Arrange
        var pptxStream = GetTestStream("001.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var slide1 = pres.Slides[0];
        var slide2 = pres.Slides[1];
        slide1.CustomData = "old-number-1";

        // Act
        slide1.Number = 2;

        // Assert
        slide1.Number.Should().Be(2);
        slide2.Number.Should().Be(1, "because the first slide was inserted to its position.");

        pres.Save();
        pres.Close();
        pres = SCPresentation.Open(pptxStream);
        slide2 = pres.Slides.First(s => s.CustomData == "old-number-1");
        slide2.Number.Should().Be(2);
    }

    [Test]
    public void Number_Setter()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slide = pres.Slides[0];

        // Act
        slide.Number = 1;

        // Assert
        slide.Number.Should().Be(1);
    }

    [Fact]
    public void GetAllTextboxes_contains_all_textboxes_withTable()
    {
        // Arrange
        var preStream = TestFiles.Presentations.pre039_stream;
        var presentation = SCPresentation.Open(preStream);
        var slide = presentation.Slides.First();

        // Act
        var textboxes = slide.GetAllTextFrames();

        // Assert
        textboxes.Count.Should().Be(11);
    }

    [Fact]
    public void GetAllTextboxes_contains_all_textboxes_withoutTable()
    {
        // Arrange
        var pptxStream = TestFiles.Presentations.pre011_dt_stream;
        var pres = SCPresentation.Open(pptxStream);
        var slide = pres.Slides.First();

        // Act
        var textFrames = slide.GetAllTextFrames();

        // Assert
        textFrames.Count.Should().Be(4);
    }

#if DEBUG

    [Fact(Skip = "In progress")]
    public void SaveAsPng_saves_slide_as_image()
    {
        // Arrange
        var pptxStream = GetTestStream("autoshape-case011_save-as-png.pptx");
        var pres = SCPresentation.Open(pptxStream);
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
        var pre = SCPresentation.Open(pptx);
        var slide = pre.Slides[0];

        // Act
        var slideHtml = slide.ToHtml();

        // Arrange
    }

#endif
}