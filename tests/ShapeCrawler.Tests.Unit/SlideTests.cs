using System.Diagnostics.CodeAnalysis;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;

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
        var pptx = TestAsset("001.pptx");
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
        var pptx = TestAsset("002.pptx");
        var pres = new Presentation(pptx);
        ISlide slideEx = pres.Slides[2];

        // Act
        bool hidden = slideEx.Hidden();

        // Assert
        hidden.Should().BeTrue();
    }

    [Test]
    public void Fill_Picture_Update_updates_background()
    {
        // Arrange
        var pre = new Presentation(TestAsset("009_table.pptx"));
        var backgroundImage = pre.Slides[0].Fill.Picture;
        var image = TestAsset("10 png image.png");
        var bytesBefore = backgroundImage.AsByteArray();

        // Act
        backgroundImage.Update(image);

        // Assert
        var bytesAfter = backgroundImage.AsByteArray();
        bytesAfter.Length.Should().NotBe(bytesBefore.Length);
    }

    [Test]
    public void Fill_SetPicture_sets_image_background()
    {
        // Arrange
        var pres = new Presentation();
        pres.Slides.AddEmptySlide(SlideLayoutType.Blank);
        var slide = pres.Slides[0];
        var image = TestAsset("10 png image.png");

        // Act
        slide.Fill.SetPicture(image);

        // Assert
        slide.Fill.Picture.Should().NotBeNull();
        pres.Validate();
    }
    
    [Test]
    public void Fill_Picture_is_null_when_slide_doesnt_have_background()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var slideFill = pres.Slide(2).Fill;

        // Act-Assert
        slideFill.Picture.Should().BeNull();
    }

    [Test]
    public void Fill_Color_gets_returns_solid_color_of_slide_background()
    {
        // Arrange
        var pres = new Presentation(TestAsset("058_bg-fill.pptx"));
        var slideFill = pres.Slide(1).Fill;
        var orange = "E6BB90";

        // Act-Assert
        slideFill.Color.Should().Be(orange);
    }

    [Test]
    public void Fill_SetColor_sets_solid_color_for_slide_background()
    {
        // Arrange
        var pres = new Presentation();
        var slideFill = pres.Slide(1).Fill;
        var green = "00ff00";

        // Act
        slideFill.SetColor(green);

        // Assert
        slideFill.Color.Should().Be(green);
    }

    [Test]
    public void Fill_SolidColor_twice_sets_fill()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var expected = "ABCDEF";

        // Act
        slide.Fill.SetColor("123456");
        slide.Fill.SetColor(expected);

        // Assert
        var actual = slide.Fill.Color;
        actual.Should().Be(expected);
    }

    [Test]
    public void CustomData_ReturnsData_WhenCustomDataWasAssigned()
    {
        // Arrange
        const string customDataString = "Test custom data";
        var originPre = new Presentation(TestAsset("001.pptx"));
        var slide = originPre.Slides.First();

        // Act
        slide.CustomData = customDataString;

        var savedPreStream = new MemoryStream();
        originPre.Save(savedPreStream);
        var savedPre = new Presentation(savedPreStream);
        var customData = savedPre.Slides.First().CustomData;

        // Assert
        customData.Should().Be(customDataString);
    }

    [Test]
    public void CustomData_PropertyIsNull_WhenTheSlideHasNotCustomData()
    {
        // Arrange
        var slide = new Presentation(TestAsset("001.pptx")).Slides.First();

        // Act
        var sldCustomData = slide.CustomData;

        // Assert
        sldCustomData.Should().BeNull();
    }

    [Test]
    public void Number_Setter_moves_slide_to_specified_number_position()
    {
        // Arrange
        var pptxStream = TestAsset("001.pptx");
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
        var pptx = TestAsset("039.pptx");
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
        var pres = new Presentation(TestAsset("011_dt.pptx"));
        var slide = pres.Slides.First();

        // Act
        var textFrames = slide.TextFrames();

        // Assert
        textFrames.Count.Should().Be(4);
    }

    [Test]
    public void Notes_Getter_returns_notes()
    {
        // Arrange
        var pptxStream = TestAsset("056_slide-notes.pptx");
        var pres = new Presentation(pptxStream);
        var slide = pres.Slides[0];

        // Act
        var notes = slide.Notes;

        // Assert
        notes.Paragraphs.Should().HaveCount(4);
        notes.Text.Should().Contain("NOTES LINE 1");
    }

    [Test]
    public void Notes_Paragraph_Text_Setter_updates_paragraph_of_note()
    {
        // Arrange
        var pres = new Presentation(TestAsset("056_slide-notes.pptx"));
        var slide = pres.Slides[0];
        var expected = string.Join(Environment.NewLine, "0", "1", "2", "3");

        // Act
        slide.Notes.Paragraphs[0].Text = "0";
        slide.Notes.Paragraphs[1].Text = "1";
        slide.Notes.Paragraphs[2].Text = "2";
        slide.Notes.Paragraphs[3].Text = "3";

        // Assert
        slide.Notes.Paragraphs.Should().HaveCount(4);
        slide.Notes.Text.Should().Be(expected);
    }

    [Test]
    public void Notes_Paragraph_Text_Setter_updates_paragraph_of_added_note()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.AddNotes(new[] { "SlideAddNotes_can_change_notes_with_many_lines" });
        var notes = slide.Notes;
        var expected = string.Join(Environment.NewLine, "1", "2", "3");

        // Act
        notes.Paragraphs.Last().Text = "1";
        notes.Paragraphs.Add();
        notes.Paragraphs.Last().Text = "2";
        notes.Paragraphs.Add();
        notes.Paragraphs.Last().Text = "3";

        // Assert
        notes.Paragraphs.Should().HaveCount(3);
        notes.Text.Should().Be(expected);
        pres.Validate();
    }

    [Test]
    public void AddNotes_with_no_notes_adds_empty_line()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];

        // Act
        slide.AddNotes(Enumerable.Empty<string>());
        var notes = slide.Notes;

        // Assert
        notes.Text.Should().Be(string.Empty);
        pres.Validate();
    }

    [Test]
    public void Notes_Getter_returns_null_if_no_notes()
    {
        // Arrange
        var pres = new Presentation(TestAsset("003.pptx"));
        var slide = pres.Slides[0];

        // Act-Assert
        slide.Notes.Should().BeNull();
    }

    [Test]
    public void Notes_Getter_returns_null_for_new_presentation()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];

        // Act-Assert
        slide.Notes.Should().BeNull();
    }

    [Test]
    public void AddNotes_adds_notes()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var expected = "SlideAddNotes_adds_notes";

        // Act
        slide.AddNotes(new[] { expected });
        var notes = slide.Notes;

        // Assert
        notes.Text.Should().Be(expected);
        pres.Validate();
    }

    [Test]
    public void AddNotes_adds_many_notes()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        var adding = new[] { "1", "2", "3" };
        var expected = string.Join(Environment.NewLine, adding);

        // Act
        slide.AddNotes(adding);

        // Assert
        slide.Notes.Text.Should().Be(expected);
        pres.Validate();
    }

    [Test]
    public void Notes_Text_Setter_updates_notes()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.AddNotes(new[] { "Starting value" });
        var notes = slide.Notes;
        var expected = "SlideAddNotes_can_change_notes";

        // Act
        notes.Text = expected;

        // Assert
        notes.Text.Should().Be(expected);
        pres.Validate();
    }
}