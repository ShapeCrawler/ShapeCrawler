using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class BulletTests : SCTest
{
    [Test]
    public void Type_Getter_returns_None_value_When_paragraph_doesnt_have_bullet()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var shape = pres.Slides[0].Shapes.GetById<IShape>(2);
        var bullet = shape.TextBox!.Paragraphs[0].Bullet;

        // Act-Assert
        bullet.Type.Should().Be(BulletType.None);
    }
    
    [Test]
    public void Type_Getter_returns_bullet_type()
    {
        // Arrange
        var pres = new Presentation(TestAsset("002.pptx"));
        var shapeList = pres.Slides[1].Shapes;
        var shape4 = shapeList.First(x => x.Id == 4);
        var shape5 = shapeList.First(x => x.Id == 5);
        var shape4Pr2Bullet = shape4.TextBox!.Paragraphs[1].Bullet;
        var shape5Pr1Bullet = shape5.TextBox!.Paragraphs[0].Bullet;
        var shape5Pr2Bullet = shape5.TextBox.Paragraphs[1].Bullet;

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
    public void Type_Setter_updates_bullet_type()
    {
        // Arrange
        var pres = new Presentation(TestAsset("autoshape-case003.pptx"));
        var shape = pres.Slides[0].Shapes.Shape<IShape>("AutoShape 1");
        var bullet = shape.TextBox!.Paragraphs[0].Bullet;

        // Act
        bullet.Type = BulletType.Character;
        bullet.Character = "*";

        // Assert
        bullet.Type.Should().Be(BulletType.Character);
        bullet.Character.Should().Be("*");

        var savedPreStream = new MemoryStream();
        pres.Save(savedPreStream);
        var newPresentation = new Presentation(savedPreStream);
        shape = newPresentation.Slides[0].Shapes.Shape<IShape>("AutoShape 1");
        bullet = shape.TextBox!.Paragraphs[0].Bullet;
        bullet.Type.Should().Be(BulletType.Character);
        bullet.Character.Should().Be("*");
    }
    
    [Test]
    public void Type_Setter_sets_Numbered_bullet_type()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = new Presentation(TestAsset("020.pptx"));
        var shape = pres.Slides[2].Shapes.First(sp => sp.Id == 7);
        shape.TextBox!.Paragraphs.Add();
        var addedParagraph = shape.TextBox.Paragraphs.Last();

        // Act
        addedParagraph.Bullet.Type = BulletType.Numbered;
        addedParagraph.Bullet.Size = 100;
        addedParagraph.Bullet.FontName = "Tahoma";

        // Assert
        addedParagraph.Bullet.Type.Should().Be(BulletType.Numbered);
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");

        pres.Save(mStream);

        pres = new Presentation(mStream);
        shape = pres.Slides[2].Shapes.First(sp => sp.Id == 7);
        addedParagraph = shape.TextBox!.Paragraphs.Last();
        addedParagraph.Bullet.Type.Should().Be(BulletType.Numbered);
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");
    }
    
    [Test]
    public void Character_Getter_returns_list_paragraph_bullet_character()
    {
        // Arrange
        var pres = new Presentation(pres =>
        {
            pres.Slide(slide =>
            {
                slide.RectangleShape(textBox => 
                {
                    textBox.Paragraph(para => {
                        para.Text("Hello, World!");
                        para.BulletedList();
                    });
                });
            });
        });
        var paragraph = pres.Slide(1).Shapes.First().TextBox!.Paragraphs.First();

        // Act-Assert
        paragraph.Bullet.Character.Should().Be("•");
    }
    
    [Test]
    public void Character_Setter_updates_bullet_character()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = new Presentation(TestAsset("020.pptx"));
        var shape = pres.Slide(3).Shapes.First(sp => sp.Id == 7);
        shape.TextBox!.Paragraphs.Add();
        var addedParagraph = shape.TextBox.Paragraphs.Last();

        // Act
        addedParagraph.Bullet.Type = BulletType.Character;
        addedParagraph.Bullet.Character = "*";
        addedParagraph.Bullet.Size = 100;
        addedParagraph.Bullet.FontName = "Tahoma";

        // Assert
        addedParagraph.Bullet.Type.Should().Be(BulletType.Character);
        addedParagraph.Bullet.Character.Should().Be("*");
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");

        pres.Save(mStream);

        pres = new Presentation(mStream);
        shape = pres.Slides[2].Shapes.First(sp => sp.Id == 7);
        addedParagraph = shape.TextBox!.Paragraphs.Last();
        addedParagraph.Bullet.Type.Should().Be(BulletType.Character);
        addedParagraph.Bullet.Character.Should().Be("*");
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");
    }
    
    [Test]
    public void FontName_Getter_returns_font_name()
    {
        // Arrange
        var pres = new Presentation(TestAsset("002.pptx"));
        var shapes = pres.Slides[1].Shapes;
        var shape3Pr1Bullet = shapes.First(x => x.Id == 3).TextBox!.Paragraphs[0].Bullet;
        var shape4Pr2Bullet = shapes.First(x => x.Id == 4).TextBox!.Paragraphs[1].Bullet;

        // Act
        var shape3BulletFontName = shape3Pr1Bullet.FontName;
        var shape4BulletFontName = shape4Pr2Bullet.FontName;

        // Assert
        shape3BulletFontName.Should().BeNull();
        shape4BulletFontName.Should().Be("Calibri");
    }
    
    [Test]
    public void ColorHex_Getter_returns_bullet_properties()
    {
        // Arrange
        var pres = new Presentation(TestAsset("002.pptx"));
        var shape = pres.Slide(2).Shapes.First(x => x.Id == 4);
        var bullet = shape.TextBox!.Paragraphs[1].Bullet;

        // Act
        var bulletColor = bullet.ColorHex;
        var bulletChar = bullet.Character;
        var bulletSize = bullet.Size;

        // Assert
        bulletColor.Should().Be("C00000");
        bulletChar.Should().Be("'");
        bulletSize.Should().Be(120);
    }
}