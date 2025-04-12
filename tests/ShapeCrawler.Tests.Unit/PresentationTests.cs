using System.Globalization;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Presentations;
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
    public void SlideWidth_Getter_returns_presentation_Slides_Width()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));

        // Act & Assert
        pres.SlideWidth.Should().Be(720);
    }

    [Test]
    public void SlideWidth_Setter_sets_presentation_Slides_Width()
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
    public void SlideHeight_Getter_returns_presentation_Slides_Height()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));

        // Act & Assert
        pres.SlideHeight.Should().Be(405);
    }

    [Test]
    public void SlideHeight_Setter_sets_presentation_Slides_Height()
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
        // Arrange
        var pres17 = new Presentation(TestAsset("017.pptx"));
        var pres16 = new Presentation(TestAsset("016.pptx"));
        var pres75 = new Presentation(TestAsset("075.pptx"));

        // Act & Assert
        pres17.Slides.Count.Should().Be(1);
        pres16.Slides.Count.Should().Be(1);
        pres75.Slides.Count.Should().Be(1);
    }

    [Test]
    public void Slides_Count()
    {
        // Arrange
        var pres = new Presentation(TestAsset("007_2 slides.pptx"));
        var removingSlide = pres.Slides[0];
        var slides = pres.Slides;

        // Act
        removingSlide.Remove();

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

        destPre.Save(savedPre);
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

        destPres.Save(savedPre);
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
        removingSlide.Remove();
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
        // TODO: Add assertion
    }

#if DEBUG
    [Test]
    [Explicit]
    public void Slides_Add()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case017_slide-number.pptx"));
        const string jsonSlide = """
                                 {
                                   "slideLayoutNumber": 1,
                                   "slideLayoutShapes": [
                                     {
                                       "name": "Title",
                                       "text": "Hello World!"
                                     }
                                   ]
                                 }
                                 """;
        // Act
        pres.Slides.AddJSON(jsonSlide);

        // Assert
        pres.Validate();
    }
#endif

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
        pres.Slides[0].Remove();
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
        pres.Save(stream);

        // Assert
        stream.Position = 0;
        var updatedPres = new Presentation(stream);
        updatedPres.Properties.Modified.Should().Be(expectedModified);
    }

    [Test]
    public void Footer_AddSlideNumber_adds_slide_number()
    {
        // Arrange
        var pres = new Presentation();

        // Act
        pres.Footer.AddSlideNumber();

        // Assert
        pres.Footer.SlideNumberAdded().Should().BeTrue();
    }

    [Test, Ignore("In Progress #540")]
    public void Footer_RemoveSlideNumber_removes_slide_number()
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
    public void Footer_SlideNumberAdded_returns_false_When_slide_number_is_not_added()
    {
        // Arrange
        var pres = new Presentation();

        // Act-Assert
        pres.Footer.SlideNumberAdded().Should().BeFalse();
    }

    [Test]
    public void Slides_Add_adds_slide()
    {
        // Arrange
        var sourceSlide = new Presentation(TestAsset("001.pptx")).Slides[0];
        var pptx = TestAsset("002.pptx");
        var destPre = new Presentation(pptx);
        var originSlidesCount = destPre.Slides.Count;
        var expectedSlidesCount = ++originSlidesCount;
        MemoryStream savedPre = new();

        // Act
        destPre.Slides.Add(sourceSlide);

        // Assert
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");

        destPre.Save(savedPre);
        destPre = new Presentation(savedPre);
        destPre.Slides.Count.Should().Be(expectedSlidesCount, "because the new slide has been added");
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
        removingSlide.Remove();

        // Assert
        pres.Slides.Should().HaveCount(expectedSlidesCount);

        pres.Save(mStream);
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
        pres.Properties.Title = "Properties_setter_sets_values";
        pres.Properties.Created = expectedCreated;

        // Assert
        pres.Properties.Title.Should().Be("Properties_setter_sets_values");
        pres.Properties.Created.Should().Be(expectedCreated);
    }

    [Test]
    public void FileProperties_getters_return_valid_values_after_saving_presentation()
    {
        // Arrange
        var pres = new Presentation();
        var expectedCreated = new DateTime(2024, 1, 2, 3, 4, 5, DateTimeKind.Local);
        var stream = new MemoryStream();

        // Act
        pres.Properties.Title = "Properties_setter_survives_round_trip";
        pres.Properties.Created = expectedCreated;
        pres.Properties.RevisionNumber = 100;
        pres.Save(stream);

        // Assert
        stream.Position = 0;
        var updatePres = new Presentation(stream);
        updatePres.Properties.Title.Should().Be("Properties_setter_survives_round_trip");
        updatePres.Properties.Created.Should().Be(expectedCreated);
        pres.Properties.RevisionNumber.Should().Be(100);
    }

    [Test]
    public void FileProperties_Modified_Getter_returns_date_of_the_last_modification()
    {
        var pres = new Presentation(TestAsset("059_crop-images.pptx"));
        var expectedModified = DateTime.Parse("2024-12-16T17:11:58Z", CultureInfo.InvariantCulture);

        // Act-Assert
        pres.Properties.Modified.Should().Be(expectedModified);
        pres.Properties.Title.Should().Be("");
        pres.Properties.RevisionNumber.Should().Be(7);
        pres.Properties.Comments.Should().BeNull();
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
        pres.Properties.Modified.Should().Be(expectedModified);
    }

    [Test]
    [Explicit("Should be fixed")]
    public void Constructor_does_not_throw_exception_When_the_specified_file_is_a_google_slide_export()
    {
        // Act
        var openingGoogleSlides = () => new Presentation(TestAsset("074 google slides.pptx"));

        // Assert
        openingGoogleSlides.Should().NotThrow();
    }

    [Test]
    public void AsMarkdown_returns_markdown_string()
    {
        // Arrange
        var pres = new Presentation(TestAsset("076 bitcoin.pptx"));
        var expectedMarkdown = StringOf("076 bitcoin.md");

        // Act & Assert
        pres.AsMarkdown().Should().Be(expectedMarkdown);
    }
}