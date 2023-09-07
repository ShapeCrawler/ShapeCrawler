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
        var pres = new SCPresentation();

        // Assert
        pres.Should().NotBeNull();
        pres.Validate();
    }

    [Test]
    public void SlideWidth_Getter_returns_presentation_Slides_Width_in_pixels()
    {
        // Arrange
        var pptx = StreamOf("009_table.pptx");
        var pres = new SCPresentation(pptx);

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
        var pres = new SCPresentation(pptx);

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
        var pres = new SCPresentation(pptx);

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
        var pres = new SCPresentation(pptx);

        // Act
        pres.SlideHeight = 700;

        // Assert
        pres.SlideHeight.Should().Be(700);
    }

    [Test]
    public void Slides_Count_returns_One_When_presentation_contains_one_slide()
    {
        // Act
        var pres17 = new SCPresentation(StreamOf("017.pptx"));
        var pres16 = new SCPresentation(StreamOf("016.pptx"));
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
        var sourceSlide = new SCPresentation(StreamOf("001.pptx")).Slides[0];
        var destPre = new SCPresentation(StreamOf("002.pptx"));
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new ();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.SaveAs(savedPre);
        destPre = new SCPresentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }
    
    [Test]
    public void Slides_Add_adds_should_copy_only_layout_of_copying_slide()
    {
        // Arrange
        var sourcePptx = StreamOf("pictures-case004.pptx");
        var destPptx = StreamOf("autoshape-grouping.pptx");
        var sourcePres = new SCPresentation(sourcePptx);
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new SCPresentation(destPptx);
        var expectedCount = destPres.Slides.Count + 1;
        MemoryStream savedPre = new ();

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Count.Should().Be(expectedCount);

        destPres.SaveAs(savedPre);
        destPres = new SCPresentation(savedPre);
        destPres.Slides.Count.Should().Be(expectedCount);
        destPres.Slides[1].SlideLayout.SlideMaster.SlideLayouts.Count.Should().Be(1);
        destPres.Validate();
    }

    [Test]
    public void Slides_Insert_inserts_specified_slide_at_the_specified_position()
    {
        // Arrange
        var sourceSlide = new SCPresentation(StreamOf("001.pptx")).Slides[0];
        string sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        var destPre = new SCPresentation(StreamOf("002.pptx"));

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
        var pres = new SCPresentation(pptxStream);
        var sectionSlides = pres.Sections[0].Slides;
        var removingSlide = sectionSlides[0];
        var mStream = new MemoryStream();

        // Act
        pres.Slides.Remove(removingSlide);

        // Assert
        sectionSlides.Count.Should().Be(0);

        pres.SaveAs(mStream);
        pres = new SCPresentation(mStream);
        sectionSlides = pres.Sections[0].Slides;
        sectionSlides.Count.Should().Be(0);
    }

    [Test]
    public void SlideMastersCount_ReturnsNumberOfMasterSlidesInThePresentation()
    {
        // Arrange
        IPresentation presentationCase1 = new SCPresentation(StreamOf("001.pptx"));
        IPresentation presentationCase2 = new SCPresentation(StreamOf("002.pptx"));

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
        var pres = new SCPresentation(pptx);

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
        var pres = new SCPresentation(pptxStream);
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
        var pres = new SCPresentation(pptxStream);
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
        var pres = new SCPresentation(pptxStream);
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
        var pres = new SCPresentation(pptxStream);
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
        var pres = new SCPresentation(pptx);
        var textBox = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 2").TextFrame!;
        textBox.Text = "Test";
            
        // Act
        pres.Save();
            
        // Assert
        pres = new SCPresentation(pptx);
        textBox = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 2").TextFrame!;
        textBox.Text.Should().Be("Test");
    }
    
    [Test]
    public void SaveAs_should_not_change_the_Original_Stream_when_it_is_saved_to_New_Stream()
    {
        // Arrange
        var originalStream = StreamOf("001.pptx");
        var pres = new SCPresentation(originalStream);
        var textBox = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextFrame;
        var originalText = textBox!.Text;
        var newStream = new MemoryStream();

        // Act
        textBox.Text = originalText + "modified";
        pres.SaveAs(newStream);
            
        pres = new SCPresentation(originalStream);
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
        var pres = new SCPresentation(pptx);
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
        var pres = new SCPresentation();
        
        // Act
        pres.HeaderAndFooter.AddSlideNumber();

        // Assert
        pres.HeaderAndFooter.SlideNumberAdded().Should().BeTrue();
    }
    
    [Test]
    public void HeaderAndFooter_RemoveSlideNumber_removes_slide_number()
    {
        // Arrange
        var pres = new SCPresentation();
        pres.HeaderAndFooter.AddSlideNumber();
        
        // Act
        pres.HeaderAndFooter.RemoveSlideNumber();
        
        // Assert
        pres.HeaderAndFooter.SlideNumberAdded().Should().BeFalse();
    }
    
    [Test]
    public void HeaderAndFooter_SlideNumberAdded_returns_false_When_slide_number_is_not_added()
    {
        // Arrange
        var pres = new SCPresentation();
        
        // Act-Assert
        pres.HeaderAndFooter.SlideNumberAdded().Should().BeFalse();
    }
}