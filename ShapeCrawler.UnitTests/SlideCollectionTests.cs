using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using Xunit;

namespace ShapeCrawler.Tests;

public class SlideCollectionTests : ShapeCrawlerTest
{
    [Fact]
    public void Count_returns_one_When_presentation_contains_one_slide()
    {
        // Act
        var pptx17 = GetTestStream("017.pptx");
        var pres17 = SCPresentation.Open(pptx17);        
        var pptx16 = GetTestStream("016.pptx");
        var pres16 = SCPresentation.Open(pptx16);
        var numberSlidesCase1 = pres17.Slides.Count;
        var numberSlidesCase2 = pres16.Slides.Count;

        // Assert
        numberSlidesCase1.Should().Be(1);
        numberSlidesCase2.Should().Be(1);
    }

    [Fact]
    public void Add_adds_slide_from_External_presentation()
    {
        // Arrange
        var pres1 = SCPresentation.Open(GetTestStream("001.pptx"));
        var sourceSlide = SCPresentation.Open(GetTestStream("001.pptx")).Slides[0];
        var pptx = GetTestStream("002.pptx");
        var destPre = SCPresentation.Open(pptx);
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new ();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.SaveAs(savedPre);
        destPre = SCPresentation.Open(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }
        
    [Fact]
    public void Add_adds_slide_from_the_Same_presentation()
    {
        // Arrange
        var pptxStream = GetTestStream("charts-case003.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var expectedSlidesCount = pres.Slides.Count + 1;
        var slideCollection = pres.Slides;
        var addingSlide = slideCollection[0];

        // Act
        pres.Slides.Add(addingSlide);

        // Assert
        pres.Slides.Count.Should().Be(expectedSlidesCount);
    }

    [Fact]
    public void Add_add_adds_New_slide()
    {
        // Arrange
        var pptx = GetTestStream("autoshape-case015.pptx");
        var pres = SCPresentation.Open(pptx);
        var layout = pres.SlideMasters[0].SlideLayouts[0]; 
        var slides = pres.Slides;

        // Act
        var addedSlide = slides.AddEmptySlide(layout);

        // Assert
        addedSlide.Should().NotBeNull();
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Fact]
    public void Slides_Insert_inserts_slide_at_the_specified_position()
    {
        // Arrange
        var pptx = GetTestStream("001.pptx");
        var sourceSlide = SCPresentation.Open(pptx).Slides[0];
        var sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        pptx = GetTestStream("002.pptx");
        var destPre = SCPresentation.Open(pptx);

        // Act
        destPre.Slides.Insert(2, sourceSlide);

        // Assert
        destPre.Slides[1].CustomData.Should().Be(sourceSlideId);
    }

    [Theory]
    [MemberData(nameof(TestCasesSlidesRemove))]
    public void Slides_Remove_removes_slide(string file, int expectedSlidesCount)
    {
        // Arrange
        var pptx = GetTestStream(file);
        var pres = SCPresentation.Open(pptx);
        var removingSlide = pres.Slides[0];
        var mStream = new MemoryStream();

        // Act
        pres.Slides.Remove(removingSlide);

        // Assert
        pres.Slides.Should().HaveCount(expectedSlidesCount);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        pres.Slides.Should().HaveCount(expectedSlidesCount);
    }
        
    public static IEnumerable<object[]> TestCasesSlidesRemove()
    {
        yield return new object[] {"007_2 slides.pptx", 1};
        yield return new object[] {"006_1 slides.pptx", 0};
    }
        
    [Fact]
    public void Slides_Remove_removes_slide_from_section()
    {
        // Arrange
        var pptxStream = GetTestStream("030.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var sectionSlides = pres.Sections[0].Slides;
        var removingSlide = sectionSlides[0];
        var mStream = new MemoryStream();

        // Act
        pres.Slides.Remove(removingSlide);

        // Assert
        sectionSlides.Count.Should().Be(0);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        sectionSlides = pres.Sections[0].Slides;
        sectionSlides.Count.Should().Be(0);
    }
}