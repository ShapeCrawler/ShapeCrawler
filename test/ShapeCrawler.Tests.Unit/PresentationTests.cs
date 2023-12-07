using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Charts;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class PresentationTests : SCTest
{
    [Test]
    public void Create_creates_a_new_presentation()
    {
        // Act
        var pres = new Presentation();

        // Assert
        pres.Should().NotBeNull();
        pres.Validate();
    }

    [Test]
    public void SlideWidth_Getter_returns_presentation_Slides_Width_in_pixels()
    {
        // Arrange
        var pptx = StreamOf("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        var slidesWidth = pres.SlideWidth;

        // Assert
        slidesWidth.Should().Be(960);
    }
    
    [Test]
    public void SlideWidth_Setter_sets_presentation_Slides_Width_in_pixels()
    {
        // Arrange
        var pptx = StreamOf("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        pres.SlideWidth = 1000;

        // Assert
        pres.SlideWidth.Should().Be(1000);
    }
        
    [Test]
    public void SlideHeight_Getter_returns_presentation_Slides_Height_in_pixels()
    {
        // Arrange
        var pptx = StreamOf("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        var slideHeight = pres.SlideHeight;

        // Assert
        slideHeight.Should().Be(540);
    }
    
    [Test]
    public void SlideHeight_Setter_sets_presentation_Slides_Height_in_pixels()
    {
        // Arrange
        var pptx = StreamOf("009_table.pptx");
        var pres = new Presentation(pptx);

        // Act
        pres.SlideHeight = 700;

        // Assert
        pres.SlideHeight.Should().Be(700);
    }

    [Test]
    public void Slides_Count_returns_One_When_presentation_contains_one_slide()
    {
        // Act
        var pres17 = new Presentation(StreamOf("017.pptx"));
        var pres16 = new Presentation(StreamOf("016.pptx"));
        var numberSlidesCase1 = pres17.Slides.Count;
        var numberSlidesCase2 = pres16.Slides.Count;

        // Assert
        numberSlidesCase1.Should().Be(1);
        numberSlidesCase2.Should().Be(1);
    }

    [Test]
    public void Slides_Add_adds_specified_slide_at_the_end_of_slide_collection()
    {
        // Arrange
        var sourceSlide = new Presentation(StreamOf("001.pptx")).Slides[0];
        var destPre = new Presentation(StreamOf("002.pptx"));
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new ();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.SaveAs(savedPre);
        destPre = new Presentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }
    
    [Test]
    public void Slides_Add_adds_should_copy_only_layout_of_copying_slide()
    {
        // Arrange
        var sourcePptx = StreamOf("pictures-case004.pptx");
        var destPptx = StreamOf("autoshape-grouping.pptx");
        var sourcePres = new Presentation(sourcePptx);
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new Presentation(destPptx);
        var expectedCount = destPres.Slides.Count + 1;
        MemoryStream savedPre = new ();

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Count.Should().Be(expectedCount);

        destPres.SaveAs(savedPre);
        destPres = new Presentation(savedPre);
        destPres.Slides.Count.Should().Be(expectedCount);
        destPres.Slides[1].SlideLayout.SlideMaster.SlideLayouts.Count.Should().Be(1);
        destPres.Validate();
    }

    [Test]
    public void Slides_Insert_inserts_specified_slide_at_the_specified_position()
    {
        // Arrange
        var sourceSlide = new Presentation(StreamOf("001.pptx")).Slides[0];
        string sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        var destPre = new Presentation(StreamOf("002.pptx"));

        // Act
        destPre.Slides.Insert(2, sourceSlide);

        // Assert
        destPre.Slides[1].CustomData.Should().Be(sourceSlideId);
    }

    [Test]
    public void Slides_Remove_removes_slide_from_section()
    {
        // Arrange
        var pptxStream = StreamOf("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var sectionSlides = pres.Sections[0].Slides;
        var removingSlide = sectionSlides[0];
        var mStream = new MemoryStream();

        // Act
        pres.Slides.Remove(removingSlide);

        // Assert
        sectionSlides.Count.Should().Be(0);

        pres.SaveAs(mStream);
        pres = new Presentation(mStream);
        sectionSlides = pres.Sections[0].Slides;
        sectionSlides.Count.Should().Be(0);
    }

    [Test]
    public void SlideMastersCount_ReturnsNumberOfMasterSlidesInThePresentation()
    {
        // Arrange
        IPresentation presentationCase1 = new Presentation(StreamOf("001.pptx"));
        IPresentation presentationCase2 = new Presentation(StreamOf("002.pptx"));

        // Act
        int slideMastersCountCase1 = presentationCase1.SlideMasters.Count;
        int slideMastersCountCase2 = presentationCase2.SlideMasters.Count;

        // Assert
        slideMastersCountCase1.Should().Be(1);
        slideMastersCountCase2.Should().Be(2);
    }

    [Test]
    public void SlideMaster_Shapes_Count_returns_number_of_master_shapes()
    {
        // Arrange
        var pptx = StreamOf("001.pptx");
        var pres = new Presentation(pptx);

        // Act
        var masterShapesCount = pres.SlideMasters[0].Shapes.Count;

        // Assert
        masterShapesCount.Should().Be(7);
    }

    [Test]
    public void Sections_Remove_removes_specified_section()
    {
        // Arrange
        var pptxStream = StreamOf("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }
        
    [Test]
    public void Sections_Remove_should_remove_section_after_Removing_Slide_from_section()
    {
        // Arrange
        var pptxStream = StreamOf("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var removingSection = pres.Sections[0];

        // Act
        pres.Slides.Remove(pres.Slides[0]);
        pres.Sections.Remove(removingSection);

        // Assert
        pres.Sections.Count.Should().Be(0);
    }
        
    [Test]
    public void Sections_Section_Slides_Count_returns_Zero_When_section_is_Empty()
    {
        // Arrange
        var pptxStream = StreamOf("008.pptx");
        var pres = new Presentation(pptxStream);
        var section = pres.Sections.GetByName("Section 2");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(0);
    }
                
    [Test]
    public void Sections_Section_Slides_Count_returns_number_of_slides_in_section()
    {
        var pptxStream = StreamOf("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var section = pres.Sections.GetByName("Section 1");

        // Act
        var slidesCount = section.Slides.Count;

        // Assert
        slidesCount.Should().Be(1);
    }

    [Test]
    public void Save_saves_presentation_opened_from_Stream_when_it_was_Saved()
    {
        // Arrange
        var pptx = StreamOf("autoshape-case003.pptx");
        var pres = new Presentation(pptx);
        var textBox = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 2").TextFrame!;
        textBox.Text = "Test";
            
        // Act
        pres.Save();
            
        // Assert
        pres = new Presentation(pptx);
        textBox = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 2").TextFrame!;
        textBox.Text.Should().Be("Test");
    }
    
    [Test]
    public void SaveAs_should_not_change_the_Original_Stream_when_it_is_saved_to_New_Stream()
    {
        // Arrange
        var originalStream = StreamOf("001.pptx");
        var pres = new Presentation(originalStream);
        var textBox = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextFrame;
        var originalText = textBox!.Text;
        var newStream = new MemoryStream();

        // Act
        textBox.Text = originalText + "modified";
        pres.SaveAs(newStream);
            
        pres = new Presentation(originalStream);
        textBox = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextFrame;
        var autoShapeText = textBox!.Text; 

        // Assert
        autoShapeText.Should().BeEquivalentTo(originalText);
    }
    
    [Test]
    public void BinaryData_returns_presentation_binary_content_After_updating_series()
    {
        // Arrange
        var pptx = TestHelper.GetStream("charts_bar-chart.pptx");
        var pres = new Presentation(pptx);
        var chart = pres.Slides[0].Shapes.GetByName<IChart>("Bar Chart 1");

        // Act
        chart.SeriesList[0].Points[0].Value = 1;
        var binaryData = pres.AsByteArray();
        
        // Assert
        binaryData.Should().NotBeNull();
    }
    
    [Test]
    public void HeaderAndFooter_AddSlideNumber_adds_slide_number()
    {
        // Arrange
        var pres = new Presentation();
        
        // Act
        pres.Footer.AddSlideNumber();

        // Assert
        pres.Footer.SlideNumberAdded().Should().BeTrue();
    }
    
    [Test, Ignore("In Progress #540")]
    public void HeaderAndFooter_RemoveSlideNumber_removes_slide_number()
    {
        // Arrange
        var pres = new Presentation();
        pres.Footer.AddSlideNumber();
        
        // Act
        pres.Footer.RemoveSlideNumber();
        
        // Assert
        pres.Footer.SlideNumberAdded().Should().BeFalse();
    }
    
    [Test]
    public void HeaderAndFooter_SlideNumberAdded_returns_false_When_slide_number_is_not_added()
    {
        // Arrange
        var pres = new Presentation();
        
        // Act-Assert
        pres.Footer.SlideNumberAdded().Should().BeFalse();
    }
    
    [Test]
    public void SaveAs_should_not_change_the_Original_Path_when_it_is_saved_to_New_Path()
    {
        // Arrange
        var originalPath = GetTestPath("001.pptx");
        var pres = new Presentation(originalPath);
        var textFrame = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextFrame;
        var originalText = textFrame!.Text;
        var newPath = Path.GetTempFileName();
        textFrame.Text = originalText + "modified";

        // Act
        pres.SaveAs(newPath);

        // Assert
        pres = new Presentation(originalPath);
        textFrame = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextFrame;
        var autoShapeText = textFrame!.Text; 
        autoShapeText.Should().BeEquivalentTo(originalText);
            
        // Clean
        File.Delete(originalPath);
        File.Delete(newPath);
    }
}