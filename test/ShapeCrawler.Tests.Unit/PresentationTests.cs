using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using ShapeCrawler.Charts;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests.Unit;

public class PresentationTests : SCTest
{
    [Fact]
    public void Close_should_not_throw_ObjectDisposedException()
    {
        // Arrange
        var pptx = GetTestStream("025_chart.pptx");
        var pres = SCPresentation.Open(pptx);
        var chart = pres.Slides[0].Shapes.GetById<IPieChart>(7);
        chart.Categories[0].Name = "new name";
        var mStream = new MemoryStream();
        pres.SaveAs(mStream);

        // Act
        Action act = () => pres.Close();

        // Assert
        act.Should().NotThrow<ObjectDisposedException>();
    }

    [Fact]
    public void Create_creates_a_new_presentation()
    {
        // Act
        var pres = SCPresentation.Create();

        // Assert
        pres.Should().NotBeNull();
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Fact]
    public void Slide_Width_returns_presentation_slides_width_in_pixels()
    {
        // Arrange
        var pres9 = SCPresentation.Open(TestHelperShared.GetStream("009_table.pptx"));
        var presentation = pres9;

        // Act
        var slideWidth = presentation.SlideWidth;

        // Assert
        slideWidth.Should().Be(960);
    }
        
    [Fact]
    public void Slide_Height_returns_presentation_slides_height_in_pixels()
    {
        // Arrange
        var pres9 = SCPresentation.Open(TestHelperShared.GetStream("009_table.pptx"));
        var presentation = pres9;

        // Act
        var slideHeight = presentation.SlideHeight;

        // Assert
        slideHeight.Should().Be(540);
    }

    [Fact]
    public void Slides_Count_returns_One_When_presentation_contains_one_slide()
    {
        // Act
        var pres17 = SCPresentation.Open(GetTestStream("017.pptx"));
        var pres16 = SCPresentation.Open(GetTestStream("016.pptx"));
        var numberSlidesCase1 = pres17.Slides.Count;
        var numberSlidesCase2 = pres16.Slides.Count;

        // Assert
        numberSlidesCase1.Should().Be(1);
        numberSlidesCase2.Should().Be(1);
    }

    [Fact]
    public void Slides_Add_adds_specified_slide_at_the_end_of_slide_collection()
    {
        // Arrange
        var sourceSlide = SCPresentation.Open(TestHelperShared.GetStream("001.pptx")).Slides[0];
        var destPre = SCPresentation.Open(GetTestStream("002.pptx"));
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
    public void Slides_Add_adds_should_copy_only_layout_of_copying_slide()
    {
        // Arrange
        var sourcePptx = GetTestStream("pictures-case004.pptx");
        var destPptx = TestHelperShared.GetStream("autoshape-grouping.pptx");
        var sourcePres = SCPresentation.Open(sourcePptx);
        var copyingSlide = sourcePres.Slides[0];
        var destPres = SCPresentation.Open(destPptx);
        var expectedCount = destPres.Slides.Count + 1;
        MemoryStream savedPre = new ();

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Count.Should().Be(expectedCount);

        destPres.SaveAs(savedPre);
        destPres = SCPresentation.Open(savedPre);
        destPres.Slides.Count.Should().Be(expectedCount);
        destPres.Slides[1].SlideLayout.SlideMaster.SlideLayouts.Count.Should().Be(1);
        var errors = PptxValidator.Validate(destPres);
        errors.Should().BeEmpty();
    }

    [Fact]
    public void Slides_Insert_inserts_specified_slide_at_the_specified_position()
    {
        // Arrange
        var sourceSlide = SCPresentation.Open(TestHelperShared.GetStream("001.pptx")).Slides[0];
        string sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        var destPre = SCPresentation.Open(GetTestStream("002.pptx"));

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

    [Fact]
    public void SlideMastersCount_ReturnsNumberOfMasterSlidesInThePresentation()
    {
        // Arrange
        IPresentation presentationCase1 = SCPresentation.Open(TestHelperShared.GetStream("001.pptx"));
        IPresentation presentationCase2 = SCPresentation.Open(GetTestStream("002.pptx"));

        // Act
        int slideMastersCountCase1 = presentationCase1.SlideMasters.Count;
        int slideMastersCountCase2 = presentationCase2.SlideMasters.Count;

        // Assert
        slideMastersCountCase1.Should().Be(1);
        slideMastersCountCase2.Should().Be(2);
    }

    [Fact]
    public void SlideMaster_Shapes_Count_returns_number_of_master_shapes()
    {
        // Arrange
        var pptx = TestHelperShared.GetStream("001.pptx");
        var pres = SCPresentation.Open(pptx);

        // Act
        var masterShapesCount = pres.SlideMasters[0].Shapes.Count;

        // Assert
        masterShapesCount.Should().Be(7);
    }

    [Fact]
    public void Sections_Remove_removes_specified_section()
    {
        // Arrange
        var pptxStream = GetTestStream("030.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }
        
    [Fact]
    public void Sections_Remove_should_remove_section_after_Removing_Slide_from_section()
    {
        // Arrange
        var pptxStream = GetTestStream("030.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Slides.Remove(pres.Slides[0]);
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }
        
    [Fact]
    public void Sections_Section_Slides_Count_returns_Zero_When_section_is_Empty()
    {
        // Arrange
        var pptxStream = GetTestStream("008.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var section = pres.Sections.GetByName("Section 2");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(0);
    }
                
    [Fact]
    public void Sections_Section_Slides_Count_returns_number_of_slides_in_section()
    {
        var pptxStream = GetTestStream("030.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var section = pres.Sections.GetByName("Section 1");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(1);
    }

    [Fact]
    public void Save_saves_presentation_opened_from_Stream_when_it_was_Saved()
    {
        // Arrange
        var pptxStream = GetTestStream("autoshape-case003.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 2").TextFrame;
        textBox.Text = "Test";
            
        // Act
        pres.Save();
        pres.Close();
            
        // Assert
        pres = SCPresentation.Open(pptxStream);
        textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 2").TextFrame;

        textBox.Text.Should().Be("Test");
    }
        
    [Fact]
    public void Close_doesnt_change_presentation_when_it_was_Not_Saved()
    {
        // Arrange
        var pptxStream = GetTestStream("autoshape-case003.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 2").TextFrame;
        textBox.Text = "Test";
            
        // Act
        pres.Close();
            
        // Assert
        pres = SCPresentation.Open(pptxStream);
        textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 2").TextFrame;

        textBox.Text.Should().NotBe("Test");
    }
        
    [Fact]
    public void SaveAs_should_not_change_the_Original_Stream_when_it_is_saved_to_New_Stream()
    {
        // Arrange
        var originalStream = TestHelperShared.GetStream("001.pptx");
        var pres = SCPresentation.Open(originalStream);
        var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var originalText = textBox!.Text;
        var newStream = new MemoryStream();

        // Act
        textBox.Text = originalText + "modified";
        pres.SaveAs(newStream);
            
        pres.Close();
        pres = SCPresentation.Open(originalStream);
        textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 3").TextFrame;
        var autoShapeText = textBox!.Text; 

        // Assert
        autoShapeText.Should().BeEquivalentTo(originalText);
    }
}