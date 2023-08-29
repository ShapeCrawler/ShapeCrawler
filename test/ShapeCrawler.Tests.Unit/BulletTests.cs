using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit;

public class BulletTests : SCTest
{
    [Test]
    public void Type_Setter_updates_bullet_type()
    {
        // Arrange
        var pptxStream = StreamOf("autoshape-case003.pptx");
        var pres = new SCPresentation(pptxStream);
        var shape = pres.Slides[0].Shapes.GetByName<IShape>("AutoShape 1");
        var bullet = shape.TextFrame!.Paragraphs[0].Bullet;

        // Act
        bullet.Type = SCBulletType.Character;
        bullet.Character = "*";

        // Assert
        bullet.Type.Should().Be(SCBulletType.Character);
        bullet.Character.Should().Be("*");

        var savedPreStream = new MemoryStream();
        pres.SaveAs(savedPreStream);
        var newPresentation = new SCPresentation(savedPreStream);
        shape = newPresentation.Slides[0].Shapes.GetByName<IShape>("AutoShape 1");
        bullet = shape.TextFrame!.Paragraphs[0].Bullet;
        bullet.Type.Should().Be(SCBulletType.Character);
        bullet.Character.Should().Be("*");
    }
        
    [Test]
    public void Character_Setter_updates_bullet_character()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pptx = StreamOf("020.pptx");
        IPresentation presentation = new SCPresentation(pptx);
        IShape placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        placeholderAutoShape.TextFrame.Paragraphs.Add();
        var addedParagraph = placeholderAutoShape.TextFrame.Paragraphs.Last();

        // Act
        addedParagraph.Bullet.Type = SCBulletType.Character;
        addedParagraph.Bullet.Character = "*";
        addedParagraph.Bullet.Size = 100;
        addedParagraph.Bullet.FontName = "Tahoma";

        // Assert
        addedParagraph.Bullet.Type.Should().Be(SCBulletType.Character);
        addedParagraph.Bullet.Character.Should().Be("*");
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");

        presentation.SaveAs(mStream);

        presentation = new SCPresentation(mStream);
        placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        addedParagraph = placeholderAutoShape.TextFrame.Paragraphs.Last();
        addedParagraph.Bullet.Type.Should().Be(SCBulletType.Character);
        addedParagraph.Bullet.Character.Should().Be("*");
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");
    }
        
    [Test]
    public void Type_Setter_sets_Numbered_bullet_type()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pptx = StreamOf("020.pptx");
        IPresentation presentation = new SCPresentation(pptx);
        IShape placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        placeholderAutoShape.TextFrame.Paragraphs.Add();
        var addedParagraph = placeholderAutoShape.TextFrame.Paragraphs.Last();

        // Act
        addedParagraph.Bullet.Type = SCBulletType.Numbered;
        addedParagraph.Bullet.Size = 100;
        addedParagraph.Bullet.FontName = "Tahoma";

        // Assert
        addedParagraph.Bullet.Type.Should().Be(SCBulletType.Numbered);
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");

        presentation.SaveAs(mStream);

        presentation = new SCPresentation(mStream);
        placeholderAutoShape = (IShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        addedParagraph = placeholderAutoShape.TextFrame.Paragraphs.Last();
        addedParagraph.Bullet.Type.Should().Be(SCBulletType.Numbered);
        addedParagraph.Bullet.Size.Should().Be(100);
        addedParagraph.Bullet.FontName.Should().Be("Tahoma");
    }
}