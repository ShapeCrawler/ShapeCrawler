using System.Diagnostics.CodeAnalysis;
using FluentAssertions;
using ImageMagick;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.DevTests;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class UserSlideTests : SCTest
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
        var pres = new Presentation(TestAsset("002.pptx"));
        var userSlide = pres.Slides[2];

        // Act
        bool hidden = userSlide.Hidden();

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
        var layout = pres.SlideMaster(1).SlideLayout("Blank");
        pres.Slides.Add(layout.Number);
        var slide = pres.Slides[0];
        var image = TestAsset("10 png image.png");

        // Act
        slide.Fill.SetPicture(image);

        // Assert
        slide.Fill.Picture.Should().NotBeNull();
        ValidatePresentation(pres);
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
        var pres = new Presentation(p => p.Slide());
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
        var pres = new Presentation(p => p.Slide());
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
    public void GetAllTextboxes_contains_all_textboxes_withTable()
    {
        // Arrange
        var pres = new Presentation(TestAsset("039.pptx"));
        var slide = pres.Slides.First();

        // Act-Assert
        slide.GetShapeTexts().Count.Should().Be(11);
    }

    [Test]
    public void GetTextBoxes_returns_all_slide_textboxes()
    {
        // Arrange
        var pres = new Presentation(TestAsset("011_dt.pptx"));
        var slide = pres.Slide(1);

        // Act & Assert
        slide.GetShapeTexts().Count.Should().Be(4);
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
        var pres = new Presentation(p => p.Slide());
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
        ValidatePresentation(pres);
    }

    [Test]
    public void AddNotes_with_no_notes_adds_empty_line()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var slide = pres.Slides[0];

        // Act
        slide.AddNotes(Enumerable.Empty<string>());
        var notes = slide.Notes;

        // Assert
        notes.Text.Should().Be(string.Empty);
        ValidatePresentation(pres);
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
        var pres = new Presentation(p => p.Slide());
        var slide = pres.Slides[0];

        // Act-Assert
        slide.Notes.Should().BeNull();
    }

    [Test]
    public void AddNotes_adds_notes()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var slide = pres.Slides[0];
        var expected = "SlideAddNotes_adds_notes";

        // Act
        slide.AddNotes(new[] { expected });
        var notes = slide.Notes;

        // Assert
        notes.Text.Should().Be(expected);
        ValidatePresentation(pres);
    }

    [Test]
    public void AddNotes_adds_many_notes()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var slide = pres.Slides[0];
        var adding = new[] { "1", "2", "3" };
        var expected = string.Join(Environment.NewLine, adding);

        // Act
        slide.AddNotes(adding);

        // Assert
        slide.Notes.Text.Should().Be(expected);
        ValidatePresentation(pres);
    }

    [Test]
    public void Notes_Text_Setter_updates_notes()
    {
        // Arrange
        var pres = new Presentation(p => p.Slide());
        var slide = pres.Slide(1);
        slide.AddNotes(["My notes"]);
        var notes = slide.Notes!;
        const string expected = "My new notes";

        // Act
        notes.SetText(expected);

        // Assert
        notes.Text.Should().Be(expected);
        ValidatePresentation(pres);
    }

    [Test]
    public void Remove_removes_slide_from_section()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case017_slide-number.pptx");
        var pres = new Presentation(pptxStream);
        var sectionSlides = pres.Sections[0].Slides;
        var removingSlide = sectionSlides[0];
        var mStream = new MemoryStream();

        // Act
        removingSlide.Remove();

        // Assert
        sectionSlides.Count.Should().Be(0);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        sectionSlides = pres.Sections[0].Slides;
        sectionSlides.Count.Should().Be(0);
    }

    [Test]
    public void SaveImageTo_saves_slide_image()
    {
        // Arrange
        var pres = new Presentation(p => { p.Slide(); });
        var slide = pres.Slide(1);
        var stream = new MemoryStream();

        // Act
        slide.SaveImageTo(stream);

        // Assert
        stream.Length.Should().BeGreaterThan(0);
    }

    [Test]
    public void SaveImageTo_saves_slide_image_with_solid_background()
    {
        // Arrange
        var pres = new Presentation(p => { p.Slide(s => { s.SolidBackground("FF0000"); }); });
        var slide = pres.Slide(1);
        var stream = new MemoryStream();

        // Act
        slide.SaveImageTo(stream);

        // Assert
        stream.Position = 0;
        using var image = SkiaSharp.SKBitmap.Decode(stream);
        var centerPixel = image.GetPixel(image.Width / 2, image.Height / 2);
        centerPixel.Red.Should().Be(255, "Red component");
        centerPixel.Green.Should().Be(0, "Green component");
        centerPixel.Blue.Should().Be(0, "Blue component");
    }

    [Test]
    [Platform(Exclude = "Linux",
        Reason =
            "ImageMagick.MagickTypeErrorException : UnableToReadFont `Arial' @ error/annotate.c/RenderFreetype/1658")]
    public void SaveImageTo_saves_slide_image_with_image_background()
    {
        // Arrange
        var image = TestImage();
        var pres = new Presentation(p => { p.Slide(s => { s.ImageBackground(image); }); });
        var slide = pres.Slide(1);
        var stream = new MemoryStream();

        // Act
        slide.SaveImageTo(stream);

        // Assert
        stream.Position = 0;
        using var bitmap = SkiaSharp.SKBitmap.Decode(stream);
        var cornerPixel = bitmap.GetPixel(10, 10); // corner to avoid "Shape" text in center

        // The test image background is peach #F5C8A8 (RGB: 245, 200, 168)
        cornerPixel.Red.Should().BeInRange(240, 250);
        cornerPixel.Green.Should().BeInRange(195, 205);
        cornerPixel.Blue.Should().BeInRange(163, 173);
    }

    [Test]
    [Platform(Exclude = "Linux", Reason = "System.InvalidOperationException : Sequence contains no elements")]
    public void SaveImageTo_saves_slide_image_with_text_box()
    {
        // ARRANGE
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.Shape(textBox =>
                {
                    textBox.X(50);
                    textBox.Y(50);
                    textBox.Width(100);
                    textBox.Height(50);
                    textBox.ShapeText("Hello, World!");
                });
            });
        });
        var slide = pres.Slide(1);
        using var stream = new MemoryStream();

        // ACT
        slide.SaveImageTo(stream);

        // ASSERT - verify the text box's background rectangle is rendered at position (50, 50) with accent1 color (4472C4)
        stream.Position = 0;
        using var bitmap = SkiaSharp.SKBitmap.Decode(stream);

        var left = ToPixels(50);
        var top = ToPixels(50);
        var right = ToPixels(50 + 100);
        var bottom = ToPixels(50 + 50);

        // Avoid the center since it's expected to contain text (and anti-aliased glyphs).
        const int inset = 10;

        AssertAccent1(bitmap.GetPixel(left + inset, top + inset));
        AssertAccent1(bitmap.GetPixel(right - inset, top + inset));
        AssertAccent1(bitmap.GetPixel(left + inset, bottom - inset));
        AssertAccent1(bitmap.GetPixel(right - inset, bottom - inset));
        return;

        void AssertAccent1(SkiaSharp.SKColor pixel)
        {
            // Accent1 theme color is #4472C4 (R=68, G=114, B=196)
            pixel.Red.Should().BeInRange(60, 76);
            pixel.Green.Should().BeInRange(106, 122);
            pixel.Blue.Should().BeInRange(188, 204);
        }

        static int ToPixels(int points) => (int)Math.Round(points * 96d / 72d, MidpointRounding.AwayFromZero);
    }

    [Test]
    [Platform(Exclude = "Linux", Reason = "Difference is 15%")]
    public Task SaveImageTo_saves_slide_with_bulleted_list()
    {
        // Arrange
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.Shape(shape =>
                {
                    shape.Paragraph(para =>
                    {
                        para.Text("Hello, World!");
                        para.BulletedList();
                    });
                });
            });
        });
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImageTo(stream);

        // Assert
        var imageBytes = stream.ToArray();
        return Verify(imageBytes, "png");
    }

    [Test]
    [Platform(Exclude = "Linux", Reason = "Difference is 15%")]
    public Task SaveImageTo_draws_Text_Shapes()
    {
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.TextShape(ts =>
                {
                    ts.Paragraph(para =>
                    {
                        para.Text("Key Principles");
                        para.Font(font =>
                        {
                            font.Size(18);
                            font.Bold();
                        });
                    });
                    ts.Paragraph(para =>
                    {
                        para.Text("Scalability first");
                        para.Font(font => font.Size(14));
                        para.Indentation(indent =>
                        {
                            indent.BeforeText(36);
                        });
                    });
                });
            });
        });
        
        using var stream = new MemoryStream();

        // Act
        pres.Slide(1).SaveImageTo(stream);

        // Assert
        var imageBytes = stream.ToArray();
        return Verify(imageBytes, "png");
    }
}