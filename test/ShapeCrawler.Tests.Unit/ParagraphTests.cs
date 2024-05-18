using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ParagraphTests : SCTest
{
    [Test]
    public void IndentLevel_Setter_sets_indent_level()
    {
        // Act
        var pres = new Presentation();
        pres.Slides[0].Shapes.AddRectangle(100,100, 500, 100);
        var addedShape = (IShape)pres.Slides[0].Shapes.Last();
        addedShape.TextFrame!.Paragraphs.Add();
        var paragraph = addedShape.TextFrame.Paragraphs.Last();
        paragraph.Text = "Test";
        
        // Act
        paragraph.IndentLevel = 2;

        // Assert
        paragraph.IndentLevel.Should().Be(2);
    }
    
    [Test]
    public void Bullet_FontName_Getter_returns_font_name()
    {
        // Arrange
        var pptx = StreamOf("002.pptx");
        var pres = new Presentation(pptx);
        var shapes = pres.Slides[1].Shapes;
        var shape3Pr1Bullet = ((IShape)shapes.First(x => x.Id == 3)).TextFrame.Paragraphs[0].Bullet;
        var shape4Pr2Bullet = ((IShape)shapes.First(x => x.Id == 4)).TextFrame.Paragraphs[1].Bullet;

        // Act
        var shape3BulletFontName = shape3Pr1Bullet.FontName;
        var shape4BulletFontName = shape4Pr2Bullet.FontName;

        // Assert
        shape3BulletFontName.Should().BeNull();
        shape4BulletFontName.Should().Be("Calibri");
    }

    [Test]
    public void Bullet_Type_Getter_returns_bullet_type()
    {
        // Arrange
        var pptx = StreamOf("002.pptx");
        var pres = new Presentation(pptx);
        var shapeList = pres.Slides[1].Shapes;
        var shape4 = shapeList.First(x => x.Id == 4);
        var shape5 = shapeList.First(x => x.Id == 5);
        var shape4Pr2Bullet = ((IShape)shape4).TextFrame.Paragraphs[1].Bullet;
        var shape5Pr1Bullet = ((IShape)shape5).TextFrame.Paragraphs[0].Bullet;
        var shape5Pr2Bullet = ((IShape)shape5).TextFrame.Paragraphs[1].Bullet;

        // Act
        var shape5Pr1BulletType = shape5Pr1Bullet.Type;
        var shape5Pr2BulletType = shape5Pr2Bullet.Type;
        var shape4Pr2BulletType = shape4Pr2Bullet.Type;

        // Assert
        shape5Pr1BulletType.Should().Be(BulletType.Numbered);
        shape5Pr2BulletType.Should().Be(BulletType.Picture);
        shape4Pr2BulletType.Should().Be(BulletType.Character);
    }

    [Test]
    public void Alignment_Setter_updates_text_alignment_of_table_Cell()
    {
        // Arrange
        var pres = new Presentation();
        pres.Slides[0].Shapes.AddTable(10, 10, 2, 2);
        var table = (ITable)pres.Slides[0].Shapes.Last();
        var cellTextFrame = table.Rows[0].Cells[0].TextFrame;
        cellTextFrame.Text = "some-text";
        var paragraph = cellTextFrame.Paragraphs[0];
        
        // Act
        paragraph.Alignment = TextAlignment.Center;
        
        // Assert
        paragraph.Alignment.Should().Be(TextAlignment.Center);
        pres.Validate();
    }

    [Test]
    public void Paragraph_Bullet_Type_Getter_returns_None_value_When_paragraph_doesnt_have_bullet()
    {
        // Arrange
        var pptx = StreamOf("001.pptx");
        var pres = new Presentation(pptx);
        var autoShape = pres.Slides[0].Shapes.GetById<IShape>(2);
        var bullet = autoShape.TextFrame.Paragraphs[0].Bullet;

        // Act
        var bulletType = bullet.Type;

        // Assert
        bulletType.Should().Be(BulletType.None);
    }

    [Test]
    public void Paragraph_BulletColorHexAndCharAndSizeProperties_ReturnCorrectValues()
    {
        // Arrange
        var pres2 = new Presentation(StreamOf("002.pptx"));
        var shapeList = pres2.Slides[1].Shapes;
        var shape4 = shapeList.First(x => x.Id == 4);
        var shape4Pr2Bullet = ((IShape)shape4).TextFrame.Paragraphs[1].Bullet;

        // Act
        var bulletColorHex = shape4Pr2Bullet.ColorHex;
        var bulletChar = shape4Pr2Bullet.Character;
        var bulletSize = shape4Pr2Bullet.Size;

        // Assert
        bulletColorHex.Should().Be("C00000");
        bulletChar.Should().Be("'");
        bulletSize.Should().Be(120);
    }
        
    [Test]
    public void Paragraph_Text_Setter_updates_paragraph_text_and_resize_shape()
    {
        // Arrange
        var pres = new Presentation(StreamOf("autoshape-case003.pptx"));
        var shape = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 4");
        var paragraph = shape.TextFrame.Paragraphs[0];
            
        // Act
        paragraph.Text = "AutoShape 4 some text";

        // Assert
        // TODO: Investigate! This is a pretty big approximation delta
        // NOTE: Doing this operation in Excel results in 45.12, so about what we had before.
        // In ShapeCrawler now, it's 50.88
        shape.Height.Should().BeApproximately(51.48m,0.01m);
        shape.Y.Should().Be(145m);
    }

    [Test]
    public void Text_Setter_sets_paragraph_text_in_New_Presentation()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddRectangle(10, 10, 10, 10);
        var addedShape = (IShape)slide.Shapes.Last();
        var paragraph = addedShape.TextFrame!.Paragraphs[0];

        // Act
        paragraph.Text = "test";

        // Assert
        paragraph.Text.Should().Be("test");
        pres.Validate();
    }
    
    [Test]
    public void Text_Setter_sets_paragraph_text_for_grouped_shape()
    {
        // Arrange
        var pres = new Presentation(StreamOf("autoshape-case003.pptx"));
        var shape = pres.Slides[0].Shapes.GetByName<IGroupShape>("Group 1").Shapes.GetByName<IShape>("Shape 1");
        var paragraph = shape.TextFrame.Paragraphs[0];
        
        // Act
        paragraph.Text = $"Safety{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}";
        
        // Assert
        paragraph.Text.Should().BeEquivalentTo($"Safety{Environment.NewLine}{Environment.NewLine}{Environment.NewLine}");
        pres.Validate();
    }
    
    [Test]
    public void Paragraph_Text_Getter_returns_paragraph_text()
    {
        // Arrange
        var textBox1 = ((IShape)new Presentation(StreamOf("008.pptx")).Slides[0].Shapes.First(sp => sp.Id == 37)).TextFrame;
        var textBox2 = ((ITable)new Presentation(StreamOf("009_table.pptx")).Slides[2].Shapes.First(sp => sp.Id == 3)).Rows[0].Cells[0]
            .TextFrame;

        // Act
        string paragraphTextCase1 = textBox1.Paragraphs[0].Text;
        string paragraphTextCase2 = textBox1.Paragraphs[1].Text;
        string paragraphTextCase3 = textBox2.Paragraphs[0].Text;

        // Assert
        paragraphTextCase1.Should().BeEquivalentTo("P1t1 P1t2");
        paragraphTextCase2.Should().BeEquivalentTo("p2");
        paragraphTextCase3.Should().BeEquivalentTo("0:0_p1_lvl1");
    }

    [Test]
    public void ReplaceText_finds_and_replaces_text()
    {
        // Arrange
        var pptxStream = StreamOf("autoshape-case003.pptx");
        var pres = new Presentation(pptxStream);
        var paragraph = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 3").TextFrame!.Paragraphs[0];
            
        // Act
        paragraph.ReplaceText("Some text", "Some text2");

        // Assert
        paragraph.Text.Should().BeEquivalentTo("Some text2");
        pres.Validate();
    }

    [Test]
    public void Paragraph_Portions_counter_returns_number_of_text_portions_in_the_paragraph()
    {
        // Arrange
        var textFrame = new Presentation(StreamOf("009_table.pptx")).Slides[2].Shapes.GetById<IShape>(2).TextFrame;

        // Act
        var portions = textFrame.Paragraphs[0].Portions;

        // Assert
        portions.Should().HaveCount(2);
    }
    
    [Test]
    public void Portions_Add()
    {
        // Arrange
        var pres = new Presentation(StreamOf("autoshape-case001.pptx"));
        var shape = pres.SlideMasters[0].Shapes.GetByName<IShape>("AutoShape 1");
        shape.TextFrame!.Paragraphs.Add();
        var paragraph = shape.TextFrame.Paragraphs.Last();
        var expectedPortionCount = paragraph.Portions.Count + 1;
        
        // Act
        paragraph.Portions.AddText(" ");
        
        // Assert
        paragraph.Portions.Count.Should().Be(expectedPortionCount);
    }
    
    [Test]
    [SlideParagraph("Case #1","autoshape-case003.pptx", 1, "AutoShape 5", 1, 1)]
    [SlideParagraph("Case #2","autoshape-case003.pptx", 1, "AutoShape 5", 2, 2)]
    public void IndentLevel_Getter_returns_indent_level(IParagraph paragraph, int expectedLevel)
    {
        // Act
        var indentLevel = paragraph.IndentLevel;

        // Assert
        indentLevel.Should().Be(expectedLevel);
    }
    
    [Test]
    [SlideShape("001.pptx", 1, "TextBox 3", TextAlignment.Center)]
    [SlideShape("001.pptx", 1, "Head 1", TextAlignment.Center)]
    public void Alignment_Getter_returns_text_alignment(IShape autoShape, TextAlignment expectedAlignment)
    {
        // Arrange
        var paragraph = autoShape.TextFrame.Paragraphs[0];

        // Act
        var textAlignment = paragraph.Alignment;

        // Assert
        textAlignment.Should().Be(expectedAlignment);
    }

    [Test]
    [TestCase("001.pptx", 1, "TextBox 4")]
    public void Alignment_Setter_updates_text_alignment(string presName, int slideNumber, string shapeName)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presName));
        var paragraph = pres.Slides[slideNumber - 1].Shapes.GetByName<IShape>(shapeName).TextFrame.Paragraphs[0];
        var mStream = new MemoryStream();
        
        // Act
        paragraph.Alignment = TextAlignment.Right;

        // Assert
        paragraph.Alignment.Should().Be(TextAlignment.Right);

        pres.SaveAs(mStream);
        pres = new Presentation(mStream);
        paragraph = pres.Slides[slideNumber - 1].Shapes.GetByName<IShape>(shapeName).TextFrame.Paragraphs[0];
        paragraph.Alignment.Should().Be(TextAlignment.Right);
    }

    [Test]
    [TestCase("002.pptx", 2, 4, 3, "Text", 1)]
    [TestCase("002.pptx", 2, 4, 3, "Text{NewLine}", 2)]
    [TestCase("002.pptx", 2, 4, 3, "Text{NewLine}Text2", 3)]
    [TestCase("002.pptx", 2, 4, 3, "Text{NewLine}Text2{NewLine}", 4)]
    public void Text_Setter_sets_paragraph_text(string presName, int slideNumber, int shapeId, int paraNumber, string paraText, int expectedPortionsCount)
    {
        // Arrange
        var pres = new Presentation(StreamOf(presName));
        var paragraph = pres.Slides[slideNumber - 1].Shapes.GetById<IShape>(shapeId).TextFrame.Paragraphs[paraNumber - 1];
        var mStream = new MemoryStream();
        paraText = paraText.Replace("{NewLine}", Environment.NewLine);

        // Act
        paragraph.Text = paraText;

        // Assert
        paragraph.Text.Should().BeEquivalentTo(paraText);
        paragraph.Portions.Count.Should().Be(expectedPortionsCount);

        pres.SaveAs(mStream);
        pres = new Presentation(mStream);
        paragraph = pres.Slides[slideNumber - 1].Shapes.GetById<IShape>(shapeId).TextFrame.Paragraphs[paraNumber - 1];
        paragraph.Text.Should().BeEquivalentTo(paraText);
        paragraph.Portions.Count.Should().Be(expectedPortionsCount);
    }
    
    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 5", 1.0)]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 4", 1.5)]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 3", 2.0)]
    public void Paragraph_Spacing_LineSpacingLines_returns_line_spacing_in_Lines(IShape shape, double expectedLines)
    {
        // Arrange
        var paragraph = shape.TextFrame!.Paragraphs[0];
            
        // Act
        var spacingLines = paragraph.Spacing.LineSpacingLines;
            
        // Assert
        spacingLines.Should().Be(expectedLines);
        paragraph.Spacing.LineSpacingPoints.Should().BeNull();
    }
    
    [Test]
    [SlideShape("autoshape-grouping.pptx", 1, "TextBox 6", 21.6)]
    public void Paragraph_Spacing_LineSpacingPoints_returns_line_spacing_in_Points(IShape shape, double expectedPoints)
    {
        // Arrange
        var paragraph = shape.TextFrame!.Paragraphs[0];
            
        // Act
        var spacingPoints = paragraph.Spacing.LineSpacingPoints;
            
        // Assert
        spacingPoints.Should().Be(expectedPoints);
        paragraph.Spacing.LineSpacingLines.Should().BeNull();
    }
}