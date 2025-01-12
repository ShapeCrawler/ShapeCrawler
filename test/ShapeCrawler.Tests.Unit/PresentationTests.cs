using System.Globalization;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Shared;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit;

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
        var pptx = TestAsset("009_table.pptx");
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
        var pptx = TestAsset("009_table.pptx");
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
        var pptx = TestAsset("009_table.pptx");
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
        var pptx = TestAsset("009_table.pptx");
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
        var pres17 = new Presentation(TestAsset("017.pptx"));
        var pres16 = new Presentation(TestAsset("016.pptx"));
        var numberSlidesCase1 = pres17.Slides.Count;
        var numberSlidesCase2 = pres16.Slides.Count;

        // Assert
        numberSlidesCase1.Should().Be(1);
        numberSlidesCase2.Should().Be(1);
    }
    
    [Test]
    public void Slides_Count()
    {
        // Arrange
        var pres = new Presentation(TestAsset("007_2 slides.pptx"));
        var removingSlide = pres.Slides[0];
        var slides = pres.Slides;

        // Act
        slides.Remove(removingSlide);
        
        // Assert
        slides.Count.Should().Be(1);
    }
    
    [Test]
    public void Slides_Add_adds_specified_slide_at_the_end_of_slide_collection()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slides[0];
        var destPre = new Presentation(TestAsset("002.pptx"));
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.SaveAs(savedPre);
        destPre = new Presentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
    }

    [Test]
    public void Slides_Add_should_copy_only_layout_of_copying_slide()
    {
        // Arrange
        var sourcePres = new Presentation(TestAsset("pictures-case004.pptx"));
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new Presentation(TestAsset("autoshape-grouping.pptx"));
        var expectedCount = destPres.Slides.Count + 1;
        MemoryStream savedPre = new();

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
    public void Slides_Add_should_copy_notes()
    {
        // Arrange
        var sourcePres = new Presentation(TestAsset("008.pptx"));
        var copyingSlide = sourcePres.Slides[0];
        var destPres = new Presentation(TestAsset("autoshape-case017_slide-number.pptx"));

        // Act
        destPres.Slides.Add(copyingSlide);

        // Assert
        destPres.Slides.Last().Notes!.Text.Should().Be("Test note");
        destPres.Validate();
    }

    [Test]
    public void Slides_AddEmptySlide()
    {
        // Arrange
        var pres = new Presentation();
        var removingSlide = pres.Slides[0];
        
        // Act
        pres.Slides.Remove(removingSlide);
        pres.Slides.AddEmptySlide(SlideLayoutType.Blank);
        
        // Assert
        pres.Slides.Count.Should().Be(1);
        pres.Validate();
    }

    [Test]
    public void Slides_Insert_inserts_specified_slide_at_the_specified_position()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slides[0];
        string sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        var destPre = new Presentation(TestAsset("002.pptx"));

        // Act
        destPre.Slides.Insert(2, sourceSlide);

        // Assert
        destPre.Slides[1].CustomData.Should().Be(sourceSlideId);
    }
    
    [Test]
    [Explicit("Should be fixed")]
    public void Slides_Insert_should_not_break_hyperlink()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case018_rotation.pptx"));
        var inserting = pres.Slide(1);

        // Act
        pres.Slides.Insert(2, inserting);
        
        // Assert
        pres.Validate();
        pres.SaveAs("output.pptx"); // uncomment for repro
        // TODO: Add assertion
    }
    
    [Test]
    public void Slides_Remove_removes_slide_from_section()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
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
        IPresentation presentationCase1 = new Presentation(TestAsset("001.pptx"));
        IPresentation presentationCase2 = new Presentation(TestAsset("002.pptx"));

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
        var pptx = TestAsset("001.pptx");
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
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
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
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
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
        var pptxStream = TestAsset("008.pptx");
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
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
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
        var pptx = TestAsset("autoshape-case003.pptx");
        var pres = new Presentation(pptx);
        var textBox = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 2").TextBox!;
        textBox.Text = "Test";

        // Act
        pres.Save();

        // Assert
        pres = new Presentation(pptx);
        textBox = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 2").TextBox!;
        textBox.Text.Should().Be("Test");
    }

    [Test]
    public void SaveAs_should_not_change_the_Original_Stream_when_it_is_saved_to_New_Stream()
    {
        // Arrange
        var originalStream = TestAsset("001.pptx");
        var pres = new Presentation(originalStream);
        var textBox = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextBox;
        var originalText = textBox!.Text;
        var newStream = new MemoryStream();

        // Act
        textBox.Text = originalText + "modified";
        pres.SaveAs(newStream);

        pres = new Presentation(originalStream);
        textBox = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextBox;
        var autoShapeText = textBox!.Text;

        // Assert
        autoShapeText.Should().BeEquivalentTo(originalText);
    }
    
    [Test]
    public void SaveAs_sets_the_creation_date()
    {
        // Arrange
        var expectedCreated = DateTime.Parse("2024-01-01T12:34:56Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedCreated);
        var stream = new MemoryStream();

        // Act
        var pres = new Presentation();
        pres.SaveAs(stream);

        // Assert
        stream.Position = 0;
        var updatedPres = new Presentation(stream);
        updatedPres.FileProperties.Created.Should().Be(expectedCreated);
    }
    
    [Test]
    public void SaveAs_sets_the_date_of_the_last_modification()
    {
        // Arrange
        var expectedCreated = DateTime.Parse("2024-01-01T12:34:56Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedCreated);
        var pres = new Presentation();
        var expectedModified = DateTime.Parse("2024-02-02T15:30:45Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedModified);
        var stream = new MemoryStream();

        // Act
        pres.SaveAs(stream);

        // Assert
        stream.Position = 0;
        var updatedPres = new Presentation(stream);
        updatedPres.FileProperties.Modified.Should().Be(expectedModified);
    } 
    
    [Test]
    public void BinaryData_returns_presentation_binary_content_After_updating_series()
    {
        // Arrange
        var pptx = TestAsset("charts_bar-chart.pptx");
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
        var textFrame = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextBox;
        var originalText = textFrame!.Text;
        var newPath = Path.GetTempFileName();
        textFrame.Text = originalText + "modified";

        // Act
        pres.SaveAs(newPath);

        // Assert
        pres = new Presentation(originalPath);
        textFrame = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 3").TextBox;
        var autoShapeText = textFrame!.Text;
        autoShapeText.Should().BeEquivalentTo(originalText);

        // Clean
        File.Delete(originalPath);
        File.Delete(newPath);
    }

    [Test]
    [Parallelizable(ParallelScope.None)]
    public void Slides_Add_adds_slide()
    {
        // Arrange
        var source = new Presentation(TestAsset("001.pptx"));
        var targetPath = GetTestPath("008.pptx");
        var target = new Presentation(targetPath);
        var copyingSlide = source.Slides[0];

        // Act
        var slideAdding = () => target.Slides.Add(copyingSlide);

        // Assert
        slideAdding.Should().NotThrow();

        // Clean
        File.Delete(targetPath);
    }
    
    [Test]
    [TestCase("007_2 slides.pptx", 1)]
    [TestCase("006_1 slides.pptx", 0)]
    public void Slides_Remove_removes_slide(string file, int expectedSlidesCount)
    {
        // Arrange
        var pres = new Presentation(TestAsset(file));
        var removingSlide = pres.Slides[0];
        var mStream = new MemoryStream();

        // Act
        pres.Slides.Remove(removingSlide);

        // Assert
        pres.Slides.Should().HaveCount(expectedSlidesCount);

        pres.SaveAs(mStream);
        pres = new Presentation(mStream);
        pres.Slides.Should().HaveCount(expectedSlidesCount);
    }
    
    [Test]
    public void Slides_Insert_inserts_slide_at_the_specified_position()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var sourceSlide = new Presentation(pptx).Slides[0];
        var sourceSlideId = Guid.NewGuid().ToString();
        sourceSlide.CustomData = sourceSlideId;
        pptx = TestAsset("002.pptx");
        var destPre = new Presentation(pptx);

        // Act
        destPre.Slides.Insert(2, sourceSlide);

        // Assert
        destPre.Slides[1].CustomData.Should().Be(sourceSlideId);
    }

    [Test]
    public void FileProperties_Title_Setter_sets_title()
    {
        // Arrange
        var pres = new Presentation();
        var expectedCreated = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Local);

        // Act
        pres.FileProperties.Title = "Properties_setter_sets_values";
        pres.FileProperties.Created = expectedCreated;
        
        // Assert
        pres.FileProperties.Title.Should().Be("Properties_setter_sets_values");
        pres.FileProperties.Created.Should().Be(expectedCreated);
    }

    [Test]
    public void FileProperties_getters_return_valid_values_after_saving_presentation()
    {
        // Arrange
        var pres = new Presentation();
        var expectedCreated = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Local);
        var stream = new MemoryStream();

        // Act
        pres.FileProperties.Title = "Properties_setter_survives_round_trip";
        pres.FileProperties.Created = expectedCreated;
        pres.FileProperties.RevisionNumber = 100;
        pres.SaveAs(stream);
        
        // Assert
        stream.Position = 0;
        var updatePres = new Presentation(stream);
        updatePres.FileProperties.Title.Should().Be("Properties_setter_survives_round_trip");
        updatePres.FileProperties.Created.Should().Be(expectedCreated);
        pres.FileProperties.RevisionNumber.Should().Be(100);
    }

    [Test]
    public void FileProperties_Modified_Getter_returns_date_of_the_last_modification()
    {
        var pres = new Presentation(TestAsset("059_crop-images.pptx"));
        var expectedModified = DateTime.Parse("2024-12-16T17:11:58Z", CultureInfo.InvariantCulture);

        // Act-Assert
        pres.FileProperties.Modified.Should().Be(expectedModified);
        pres.FileProperties.Title.Should().Be("");
        pres.FileProperties.RevisionNumber.Should().Be(7);
        pres.FileProperties.Comments.Should().BeNull();
    }
    
    [Test]
    public void Non_parameter_constructor_sets_the_date_of_the_last_modification()
    {
        // Arrange
        var expectedModified = DateTime.Parse("2024-01-01T12:34:56Z", CultureInfo.InvariantCulture);
        SCSettings.TimeProvider = new FakeTimeProvider(expectedModified);

        // Act
        var pres = new Presentation();

        // Assert
        pres.FileProperties.Modified.Should().Be(expectedModified);
    }
}