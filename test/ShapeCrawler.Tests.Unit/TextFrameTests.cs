using System.Diagnostics.CodeAnalysis;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;

// ReSharper disable All
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit.xUnit
{
    [SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
    public class TextFrameTests : SCTest
    {
        [Test]
        public void Text_Getter_returns_text_of_table_Cell()
        {
            // Arrange
            var pptx8 = GetInputStream("008.pptx");
            var pres8 = SCPresentation.Open(pptx8);
            var pptx1 = GetInputStream("001.pptx");
            var pres1 = SCPresentation.Open(pptx1);
            var pptx9 = GetInputStream("009_table.pptx");
            var pres9 = SCPresentation.Open(pptx9);
            var textFrame1 = ((IAutoShape)SCPresentation.Open(GetInputStream("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3)).TextFrame;
            var textFrame2 = ((ITable)SCPresentation.Open(GetInputStream("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame;
            var textFrame3 = ((ITable)SCPresentation.Open(GetInputStream("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
                .TextFrame;
            
            // Act
            var text1 = textFrame1.Text;
            var text2 = textFrame2.Text;
            var text3 = textFrame3.Text;

            // Act
            text1.Should().NotBeEmpty();
            text2.Should().BeEquivalentTo("id3");
            text3.Should().BeEquivalentTo($"0:0_p1_lvl1{Environment.NewLine}0:0_p2_lvl2");
        }

        [Test]
        public void Text_Getter_returns_text_from_New_Slide()
        {
            // Arrange
            var pptx = GetInputStream("031.pptx");
            var pres = SCPresentation.Open(pptx);
            var layout = pres.SlideMasters[0].SlideLayouts[0];

            // Act
            pres.Slides.AddEmptySlide(layout);
            var newSlide = pres.Slides.Last();
            var textFrame = newSlide.Shapes.GetByName<IAutoShape>("Holder 5").TextFrame;
            var text = textFrame.Text;
            
            // Assert
            text.Should().BeEquivalentTo("");
        }

        [Test]
        public void Text_Setter_can_update_content_multiple_times()
        {
            // Arrange
            var pptx = GetInputStream("autoshape-case005_text-frame.pptx");
            var pres = SCPresentation.Open(pptx);
            var textFrame = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1").TextFrame;
            var modifiedPres = new MemoryStream();

            // Act
            textFrame.Text = textFrame.Text.Replace("{{replace_this}}", "confirm this");
            textFrame.Text = textFrame.Text.Replace("{{replace_that}}", "confirm that");

            // Assert
            pres.SaveAs(modifiedPres);
            pres.Close();
            pres = SCPresentation.Open(modifiedPres);
            textFrame = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1").TextFrame;
            textFrame.Text.Should().ContainAll("confirm this", "confirm that");
        }

        [Test]
        public void Text_Setter_updates_text_box_content_and_Reduces_font_size_When_text_is_Overflow()
        {
            // Arrange
            var pptxStream = GetInputStream("001.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var textBox = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 8");
            var textFrame = textBox.TextFrame;
            var fontSizeBefore = textFrame.Paragraphs[0].Portions[0].Font.Size;
            var newText = "Shrink text on overflow";

            // Act
            textFrame.Text = newText;

            // Assert
            textFrame.Text.Should().BeEquivalentTo(newText);
            textFrame.Paragraphs[0].Portions[0].Font.Size.Should().Be(8);
        }
        
        [Test]
        public void Text_Setter_resizes_shape_to_fit_text()
        {
            // Arrange
            var pptxStream = GetInputStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 4");
            var textFrame = shape.TextFrame;

            // Act
            textFrame.Text = "AutoShape 4 some text";

            // Assert
            shape.Height.Should().Be(46);
            shape.Y.Should().Be(152);
            var errors = PptxValidator.Validate(shape.SlideStructure.Presentation);
            errors.Should().BeEmpty();
        }
        
        [Test]
        public void Text_Setter_sets_text_for_New_Shape()
        {
            // Arrange
            var pres = SCPresentation.Create();
            var shapes = pres.Slides[0].Shapes;
            var autoShape = shapes.AddRectangle( 50, 60, 100, 70);
            var textFrame = autoShape.TextFrame!;
            
            // Act
            textFrame.Text = "Test";
    
            // Assert
            textFrame.Text.Should().Be("Test");
            var errors = PptxValidator.Validate(pres);
            errors.Should().BeEmpty();
        }

        [Test]
        public void AutofitType_Setter_resizes_width()
        {
            // Arrange
            var pptxStream = GetInputStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 6");
            var textFrame = shape.TextFrame!;

            // Act
            textFrame.AutofitType = SCAutofitType.Resize;

            // Assert
            shape.Width.Should().Be(107);
            var errors = PptxValidator.Validate(pres);
            errors.Should().BeEmpty();
        }

        [Test]
        public void AutofitType_Setter_updates_height()
        {
            // Arrange
            var pptxStream = GetInputStream("autoshape-case003.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 7");
            var textFrame = shape.TextFrame!;

            // Act
            textFrame.AutofitType = SCAutofitType.Resize;

            // Assert
            shape.Height.Should().Be(35);
            var errors = PptxValidator.Validate(pres);
            errors.Should().BeEmpty();
        }
        
        [Test]
        public void AutofitType_Getter_returns_text_autofit_type()
        {
            // Arrange
            var pptx = GetInputStream("001.pptx");
            var pres = SCPresentation.Open(pptx);
            var autoShape = pres.Slides[0].Shapes.GetById<IAutoShape>(9);
            var textBox = autoShape.TextFrame;

            // Act
            var autofitType = textBox.AutofitType;

            // Assert
            autofitType.Should().Be(SCAutofitType.Shrink);
        }

        [Test]
        public void Shape_IsAutoShape()
        {
            // Arrange
            var pres8 = SCPresentation.Open(GetInputStream("008.pptx"));
            var pres21 = SCPresentation.Open(GetInputStream("021.pptx"));
            IShape shapeCase1 = SCPresentation.Open(GetInputStream("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
            IShape shapeCase2 = SCPresentation.Open(GetInputStream("021.pptx")).Slides[3].Shapes.First(sp => sp.Id == 2);
            IShape shapeCase3 = SCPresentation.Open(GetInputStream("011_dt.pptx")).Slides[0].Shapes.First(sp => sp.Id == 54275);

            // Act
            var autoShapeCase1 = shapeCase1 as IAutoShape;
            var autoShapeCase2 = shapeCase2 as IAutoShape;
            var autoShapeCase3 = shapeCase3 as IAutoShape;

            // Assert
            autoShapeCase1.Should().NotBeNull();
            autoShapeCase2.Should().NotBeNull();
            autoShapeCase3.Should().NotBeNull();
        }

        [Test]
        public void Paragraphs_Add_adds_new_text_paragraph_at_the_end_And_returns_added_paragraph()
        {
            // Arrange
            const string TEST_TEXT = "ParagraphsAdd";
            var mStream = new MemoryStream();
            var pres = SCPresentation.Open(GetInputStream("001.pptx"));
            var textFrame = ((IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 4)).TextFrame;
            int originParagraphsCount = textFrame.Paragraphs.Count;

            // Act
            var addedPara = textFrame.Paragraphs.Add();
            addedPara.Text = TEST_TEXT;

            // Assert
            var lastPara = textFrame.Paragraphs.Last(); 
            lastPara.Text.Should().BeEquivalentTo(TEST_TEXT);
            textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);

            pres.SaveAs(mStream);
            pres = SCPresentation.Open(mStream);
            textFrame = ((IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 4)).TextFrame;
            textFrame.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
            textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);
        }

        [Test]
        public void Paragraphs_Add_adds_paragraph()
        {
            // Arrange
            var pptxStream = GetInputStream("autoshape-case007.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var paragraphs = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1").TextFrame.Paragraphs;
            
            // Act
            paragraphs.Add();
            
            // Assert
            paragraphs.Should().HaveCount(6);
        }

        [Test]
        public void Paragraphs_Add_adds_new_text_paragraph_at_the_end_And_returns_added_paragraph_When_it_has_been_added_after_text_frame_changed()
        {
            var pres = SCPresentation.Open(GetInputStream("001.pptx"));
            var autoShape = (IAutoShape)pres.Slides[0].Shapes.First(sp => sp.Id == 3);
            var textBox = autoShape.TextFrame;
            var paragraphs = textBox.Paragraphs;
            var paragraph = textBox.Paragraphs.First();

            // Act
            textBox.Text = "A new text";
            var newParagraph = paragraphs.Add();

            // Assert
            newParagraph.Should().NotBeNull();
        }

        [Test]
        public void CanTextChange_returns_false()
        {
            // Arrange
            var pptxStream = GetInputStream("autoshape-case006_field.pptx");
            var pres = SCPresentation.Open(pptxStream);
            var textFrame = pres.Slides[0].Shapes.GetByName<IAutoShape>("Field 1").TextFrame;
            
            // Act
            var canTextChange = textFrame.CanChangeText();
            
            // Assert
            canTextChange.Should().BeFalse();
        }
    }
}