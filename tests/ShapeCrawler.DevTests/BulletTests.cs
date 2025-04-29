using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class BulletTests : SCTest
{
    [Test]
    public void Type_Setter_updates_bullet_type()
    {
        // Arrange
        var pptxStream = TestAsset("autoshape-case003.pptx");
        var pres = new Presentation(pptxStream);
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
    public void Character_Setter_updates_bullet_character()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pptx = TestAsset("020.pptx");
        IPresentation presentation = new Presentation(pptx);
        IShape placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        placeholderAutoShape.TextBox.Paragraphs.Add();
        var addedParagraph = placeholderAutoShape.TextBox.Paragraphs.Last();

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

        presentation.Save(mStream);

        presentation = new Presentation(mStream);
        placeholderAutoShape = presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        addedParagraph = placeholderAutoShape.TextBox.Paragraphs.Last();
        addedParagraph.Bullet.Type.Should().Be(BulletType.Character);
        addedParagraph.Bullet.Character.Should().Be("*");
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");
    }
        
    [Test]
    public void Type_Setter_sets_Numbered_bullet_type()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pptx = TestAsset("020.pptx");
        IPresentation presentation = new Presentation(pptx);
        IShape placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        placeholderAutoShape.TextBox.Paragraphs.Add();
        var addedParagraph = placeholderAutoShape.TextBox.Paragraphs.Last();

        // Act
        addedParagraph.Bullet.Type = BulletType.Numbered;
        addedParagraph.Bullet.Size = 100;
        addedParagraph.Bullet.FontName = "Tahoma";

        // Assert
        addedParagraph.Bullet.Type.Should().Be(BulletType.Numbered);
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");

        presentation.Save(mStream);

        presentation = new Presentation(mStream);
        placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        addedParagraph = placeholderAutoShape.TextBox.Paragraphs.Last();
        addedParagraph.Bullet.Type.Should().Be(BulletType.Numbered);
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");
    }
}