using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;

namespace ShapeCrawler.Tests.Unit
{
    public class TextBoxTests : SCTest
    {
        [Test]
        public void Text_Getter_returns_text_of_table_Cell()
        {
            // Arrange
            var pptx8 = StreamOf("008.pptx");
            var pres8 = new Presentation(pptx8);
            var pptx1 = StreamOf("001.pptx");
            var pres1 = new Presentation(pptx1);
            var pptx9 = StreamOf("009_table.pptx");
            var pres9 = new Presentation(pptx9);
            var textFrame1 = ((IShape)new Presentation(StreamOf("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3))
                .TextBox;
            var textFrame2 = ((ITable)new Presentation(StreamOf("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 3))
                .Rows[0].Cells[0]
                .TextBox;
            var textFrame3 =
                ((ITable)new Presentation(StreamOf("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0]
                .Cells[0]
                .TextBox;

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
            var pptx = StreamOf("031.pptx");
            var pres = new Presentation(pptx);
            var layout = pres.SlideMasters[0].SlideLayouts[0];

            // Act
            pres.Slides.AddEmptySlide(layout);
            var newSlide = pres.Slides.Last();
            var textFrame = newSlide.Shapes.GetByName<IShape>("Holder 5").TextBox;
            var text = textFrame.Text;

            // Assert
            text.Should().BeEquivalentTo("");
        }

        [Test]
        public void Text_Setter_can_update_content_multiple_times()
        {
            // Arrange
            var pres = new Presentation(StreamOf("autoshape-case005_text-frame.pptx"));
            var textFrame = pres.Slides[0].Shape("TextBox 1").TextBox;
            var modifiedPres = new MemoryStream();

            // Act
            var newText = textFrame.Text.Replace("{{replace_this}}", "confirm this");
            textFrame.Text = newText;
            newText = textFrame.Text.Replace("{{replace_that}}", "confirm that");
            textFrame.Text = newText;

            // Assert
            pres.SaveAs(modifiedPres);
            pres = new Presentation(modifiedPres);
            textFrame = pres.Slides[0].Shapes.GetByName<IShape>("TextBox 1").TextBox;
            textFrame.Text.Should().Contain("confirm this");
            textFrame.Text.Should().Contain("confirm that");
        }

        [Test]
        public void Text_Setter_updates_text_box_content_and_Reduces_font_size_When_text_is_Overflow()
        {
            // Arrange
            var pres = new Presentation(StreamOf("001.pptx"));
            var textFrame = pres.Slide(1).Shape("TextBox 8").TextBox;
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
            var pres = new Presentation(StreamOf("autoshape-case003.pptx"));
            var shape = pres.Slide(1).Shape("AutoShape 4");
            var textBox = shape.TextBox;

            // Act
            textBox.Text = "AutoShape 4 some text";

            // Assert
            shape.Height.Should().BeApproximately(51.48m, 0.01m);
            shape.Y.Should().Be(149m);
            pres.Validate();
        }

        [Test]
        public void Text_Setter_resizes_shape_to_fit_multi_paragraph_text()
        {
            // Arrange
            var pres = new Presentation(StreamOf("autoshape-case003.pptx"));
            var shape = pres.Slide(1).Shape("AutoShape 4");
            var textBox = shape.TextBox;

            for (int i = 1; i < 10; i++)
            {
                // Act
                textBox.Paragraphs.Add();
                textBox.Paragraphs.Last().Text = "AutoShape 4 some text";

                // Assert
                shape.Height.Should().BeApproximately((i * 22) + 40.48m, 0.01m);
                shape.Y.Should().Be(147m - (i * 7));
                pres.Validate();
            }
        }

        [Test]
        public void Text_Setter_sets_text_for_New_Shape()
        {
            // Arrange
            var pres = new Presentation();
            var shapes = pres.Slides[0].Shapes;
            shapes.AddShape(50, 60, 100, 70);
            var textFrame = shapes.Last().TextBox;

            // Act
            textFrame.Text = "Test";

            // Assert
            textFrame.Text.Should().Be("Test");
            pres.Validate();
        }

        [Test]
        public void AutofitType_Setter_resizes_width()
        {
            // Arrange
            var pres = new Presentation(StreamOf("autoshape-case003.pptx"));
            var shape = pres.Slide(1).Shape("AutoShape 6");
            var textFrame = shape.TextBox!;

            // Act
            textFrame.AutofitType = AutofitType.Resize;

            // Assert
            shape.Width.Should().BeApproximately(136.89m, 0.01m);
            pres.Validate();
        }

        [Test]
        public void AutofitType_Setter_updates_height()
        {
            // Arrange
            var pres = new Presentation(StreamOf("autoshape-case003.pptx"));
            var shape = pres.Slide(1).Shape("AutoShape 7");
            var textBox = shape.TextBox!;

            // Act
            textBox.AutofitType = AutofitType.Resize;

            // Assert
            shape.Height.Should().BeApproximately(40.48m, 0.01m);
            pres.Validate();
        }

        [Test]
        public void AutofitType_Getter_returns_text_autofit_type()
        {
            // Arrange
            var pptx = StreamOf("001.pptx");
            var pres = new Presentation(pptx);
            var autoShape = pres.Slides[0].Shapes.GetById<IShape>(9);
            var textBox = autoShape.TextBox;

            // Act
            var autofitType = textBox.AutofitType;

            // Assert
            autofitType.Should().Be(AutofitType.Shrink);
        }

        [Test]
        public void Shape_IsAutoShape()
        {
            // Arrange
            var pres8 = new Presentation(StreamOf("008.pptx"));
            var pres21 = new Presentation(StreamOf("021.pptx"));
            IShape shapeCase1 = new Presentation(StreamOf("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 3);
            IShape shapeCase2 = new Presentation(StreamOf("021.pptx")).Slides[3].Shapes.First(sp => sp.Id == 2);
            IShape shapeCase3 = new Presentation(StreamOf("011_dt.pptx")).Slides[0].Shapes.First(sp => sp.Id == 54275);

            // Act
            var autoShapeCase1 = shapeCase1 as IShape;
            var autoShapeCase2 = shapeCase2 as IShape;
            var autoShapeCase3 = shapeCase3 as IShape;

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
            var pres = new Presentation(StreamOf("001.pptx"));
            var textFrame = ((IShape)pres.Slides[0].Shapes.First(sp => sp.Id == 4)).TextBox;
            int originParagraphsCount = textFrame.Paragraphs.Count;

            // Act
            textFrame.Paragraphs.Add();
            var addedPara = textFrame.Paragraphs.Last();
            addedPara.Text = TEST_TEXT;

            // Assert
            var lastPara = textFrame.Paragraphs.Last();
            lastPara.Text.Should().BeEquivalentTo(TEST_TEXT);
            textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);

            pres.SaveAs(mStream);
            pres = new Presentation(mStream);
            textFrame = ((IShape)pres.Slides[0].Shapes.First(sp => sp.Id == 4)).TextBox;
            textFrame.Paragraphs.Last().Text.Should().BeEquivalentTo(TEST_TEXT);
            textFrame.Paragraphs.Should().HaveCountGreaterThan(originParagraphsCount);
        }

        [Test]
        public void Paragraphs_Add_adds_paragraph()
        {
            // Arrange
            var pptxStream = StreamOf("autoshape-case007.pptx");
            var pres = new Presentation(pptxStream);
            var paragraphs = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 1").TextBox.Paragraphs;

            // Act
            paragraphs.Add();

            // Assert
            paragraphs.Should().HaveCount(6);
        }

        [Test]
        public void
            Paragraphs_Add_adds_new_text_paragraph_at_the_end_And_returns_added_paragraph_When_it_has_been_added_after_text_frame_changed()
        {
            var pres = new Presentation(StreamOf("001.pptx"));
            var autoShape = (IShape)pres.Slides[0].Shapes.First(sp => sp.Id == 3);
            var textBox = autoShape.TextBox;
            var paragraphs = textBox.Paragraphs;
            var paragraph = textBox.Paragraphs.First();

            // Act
            textBox.Text = "A new text";
            paragraphs.Add();
            var addedParagraph = paragraphs.Last();

            // Assert
            addedParagraph.Should().NotBeNull();
        }

        [Test]
        [TestCase("autoshape-case003.pptx", 1, "AutoShape 7")]
        [TestCase("001.pptx", 1, "Head 1")]
        [TestCase("autoshape-case014.pptx", 1, "Content Placeholder 1")]
        public void AutofitType_Setter_sets_autofit_type(string file, int slideNumber, string shapeName)
        {
            // Arrange
            var pres = new Presentation(StreamOf(file));
            var shape = pres.Slides[slideNumber - 1].Shapes.GetByName(shapeName);
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox!;

            // Act
            textFrame.AutofitType = AutofitType.Resize;

            // Assert
            textFrame.AutofitType.Should().Be(AutofitType.Resize);
            pres.Validate();
        }

        [Test]
        public void Text_Setter_sets_long_text()
        {
            // Arrange
            var pres = new Presentation(StreamOf("autoshape-case013.pptx"));
            var shape = pres.Slide(1).Shape("AutoShape 1");

            // Act
            var text = shape.TextBox.Text;
            shape.TextBox.Text = "Some sentence. Some sentence";

            // Assert
            shape.Height.Should().BeApproximately(93.48m, 0.01m);
        }

        [Test]
        [SlideShape("009_table.pptx", 4, 2, "Title text")]
        [SlideShape("001.pptx", 1, 5, " id5-Text1")]
        [SlideShape("019.pptx", 1, 2, "1")]
        [SlideShape("014.pptx", 2, 5, "Test subtitle")]
        [SlideShape("011_dt.pptx", 1, 54275, "Jan 2018")]
        [SlideShape("021.pptx", 4, 2, "test footer")]
        [SlideShape("012_title-placeholder.pptx", 1, 2, "Test title text")]
        [SlideShape("012_title-placeholder.pptx", 1, 3, "P1 P2")]
        public void Text_Getter_returns_text(IShape shape, string expectedText)
        {
            // Arrange
            var textFrame = ((IShape)shape).TextBox;

            // Act
            var text = textFrame.Text;

            // Assert
            text.Should().BeEquivalentTo(expectedText);
        }

        [Test]
		[SlideShape("014.pptx", 2, 5, TextVerticalAlignment.Middle)] 
		public void VerticalAlignment_Getter_returns_vertical_alignment(IShape shape, TextVerticalAlignment expectedVAlignment)
        {
            // Arrange
            var textBox = shape.TextBox;

            // Act-Assert
            textBox.VerticalAlignment.Should().Be(expectedVAlignment);
        }

        [Test]
        [SlideShape("001.pptx", 1, 6, $"id6-Text1#NewLine#Text2")]
        [SlideShape("014.pptx", 1, 61, $"test1#NewLine#test2#NewLine#test3#NewLine#test4#NewLine#test5")]
        [SlideShape("011_dt.pptx", 1, 2, $"P1#NewLine#")]
        public void Text_Getter_returns_text_with_New_Line(IShape shape, string expectedText)
        {
            // Arrange
            expectedText = expectedText.Replace("#NewLine#", Environment.NewLine);
            var textFrame = shape.TextBox;

            // Act
            var text = textFrame.Text;

            // Assert
            text.Should().BeEquivalentTo(expectedText);
        }

        [Test]
        [TestCase("001.pptx", 1, "TextBox 2")]
        [TestCase("020.pptx", 3, "TextBox 7")]
        [TestCase("001.pptx", 2, "Header 1")]
        [TestCase("autoshape-case004_subtitle.pptx", 1, "Subtitle 1")]
        [TestCase("autoshape-case008_text-frame.pptx", 1, "AutoShape 1")]
        public void Text_Setter_updates_content(string presName, int slideNumber, string shapeName)
        {
            // Arrange
            var pres = new Presentation(StreamOf(presName));
            var textFrame = pres.Slides[slideNumber - 1].Shapes.GetByName<IShape>(shapeName).TextBox;
            var mStream = new MemoryStream();

            // Act
            textFrame.Text = "Test";

            // Assert
            textFrame.Text.Should().BeEquivalentTo("Test");
            textFrame.Paragraphs.Should().HaveCount(1);

            pres.SaveAs(mStream);
            pres = new Presentation(mStream);
            textFrame = pres.Slides[slideNumber - 1].Shapes.GetByName<IShape>(shapeName).TextBox;
            textFrame.Text.Should().BeEquivalentTo("Test");
            textFrame.Paragraphs.Should().HaveCount(1);
        }

        [Test]
        [SlideShape("autoshape-case012.pptx", 1, "Shape 1")]
        public void Text_Setter(IShape shape)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox;

            // Act
            var text = textFrame.Text;
            textFrame.Text = "some text";

            // Assert
            textFrame.Text.Should().BeEquivalentTo("some text");
        }

        [Test]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 6", false)]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", true)]
        [SlideShape("autoshape-case013.pptx", 1, "AutoShape 1", true)]
        public void TextWrapped_Getter_returns_value_indicating_whether_text_is_wrapped_in_shape(IShape shape,
            bool isTextWrapped)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox!;

            // Act
            var textWrapped = textFrame.TextWrapped;

            // Assert
            textWrapped.Should().Be(isTextWrapped);
        }

        [Test]
        [SlideShape("009_table.pptx", 3, 2, 1)]
        [SlideShape("020.pptx", 3, 8, 2)]
        [SlideShape("001.pptx", 2, 2, 1)]
        public void Paragraphs_Count_returns_number_of_paragraphs_in_the_text_box(IShape shape,
            int expectedParagraphsCount)
        {
            // Arrange
            var textFrame = ((IShape)shape).TextBox;

            // Act
            var paragraphsCount = textFrame.Paragraphs.Count;

            // Assert
            paragraphsCount.Should().Be(expectedParagraphsCount);
        }

        [Test]
        public void Paragraphs_Count_returns_number_of_paragraphs_in_the_table_cell_text_box()
        {
            // Arrange
            var pres = new Presentation(StreamOf("009_table.pptx"));
            var textFrame = pres.Slides[2].Shapes.GetById<ITable>(3).Rows[0].Cells[0].TextBox;

            // Act
            var paragraphsCount = textFrame.Paragraphs.Count;

            // Assert
            paragraphsCount.Should().Be(2);
        }

        [Test]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 0.25)]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 3", 0.30)]
        public void LeftMargin_getter_returns_left_margin_of_text_frame_in_centimeters(IShape shape,
            double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox;

            // Act
            var leftMargin = textFrame.LeftMargin;

            // Assert
            leftMargin.Should().Be((decimal)expectedMargin);
        }

        [Test]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2")]
        public void LeftMargin_setter_sets_left_margin_of_text_frame_in_centimeters(IShape shape)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox;

            // Act
            textFrame.LeftMargin = 0.5m;

            // Assert
            textFrame.LeftMargin.Should().Be(0.5m);
        }

        [Test]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 0.25)]
        public void RightMargin_getter_returns_right_margin_of_text_frame_in_centimeters(IShape shape,
            double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox;

            // Act
            var rightMargin = textFrame.RightMargin;

            // Assert
            rightMargin.Should().Be((decimal)expectedMargin);
        }

        [Test]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 0.13)]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 3", 0.14)]
        public void TopMargin_getter_returns_top_margin_of_text_frame_in_centimeters(IShape shape,
            double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox;

            // Act
            var topMargin = textFrame.TopMargin;

            // Assert
            topMargin.Should().Be((decimal)expectedMargin);
        }

        [Test]
        [SlideShape("autoshape-case003.pptx", 1, "AutoShape 2", 0.13)]
        public void BottomMargin_getter_returns_bottom_margin_of_text_frame_in_centimeters(IShape shape,
            double expectedMargin)
        {
            // Arrange
            var autoShape = (IShape)shape;
            var textFrame = autoShape.TextBox;

            // Act
            var bottomMargin = textFrame.BottomMargin;

            // Assert
            bottomMargin.Should().Be((decimal)expectedMargin);
        }

		[Test]
		[TestCase("001.pptx", 1, "TextBox 4")]
		public void VerticalAlignment_Setter_updates_text_vertical_alignment(string presName, int slideNumber, string shapeName)
		{
			// Arrange
			var pres = new Presentation(StreamOf(presName));
			var textbox = pres.Slides[slideNumber - 1].Shapes.GetByName<IShape>(shapeName).TextBox;
			var mStream = new MemoryStream();

			// Act
			textbox.VerticalAlignment = TextVerticalAlignment.Bottom;

			// Assert
			textbox.VerticalAlignment.Should().Be(TextVerticalAlignment.Bottom);

			pres.SaveAs(mStream);
			pres = new Presentation(mStream);
			textbox = pres.Slides[slideNumber - 1].Shapes.GetByName<IShape>(shapeName).TextBox;
			textbox.VerticalAlignment.Should().Be(TextVerticalAlignment.Bottom);
		}

		[Test]
        [TestCase("054_get_shape_xpath.pptx", 1, "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]")]
        [TestCase("054_get_shape_xpath.pptx", 2, "/p:sld[1]/p:cSld[1]/p:spTree[1]/p:sp[1]/p:txBody[1]")]
        public void SDKXPath_returns_xpath_of_undelying_txBody_element(string presentationName, int slideNumber, string expectedXPath)
        {
            // Arrange
            var pres = new Presentation(StreamOf(presentationName));
            var textFrame = pres.Slides[slideNumber - 1].TextFrames().First();

            // Act
            var sdkXPath = textFrame.SdkXPath;

            // Assert
            sdkXPath.Should().Be(expectedXPath);
        }
    }
}