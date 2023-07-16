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
        var pptxStream = GetInputStream("autoshape-case003.pptx");
        var pres = SCPresentation.Open(pptxStream);
        var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
        var bullet = shape.TextFrame!.Paragraphs[0].Bullet;

        // Act
        bullet.Type = SCBulletType.Character;
        bullet.Character = "*";

        // Assert
        bullet.Type.Should().Be(SCBulletType.Character);
        bullet.Character.Should().Be("*");

        var savedPreStream = new MemoryStream();
        pres.SaveAs(savedPreStream);
        var newPresentation = SCPresentation.Open(savedPreStream);
        shape = newPresentation.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
        bullet = shape.TextFrame!.Paragraphs[0].Bullet;
        bullet.Type.Should().Be(SCBulletType.Character);
        bullet.Character.Should().Be("*");
    }
        
    [Test]
    public void Character_Setter_updates_bullet_character()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pptx = GetInputStream("020.pptx");
        IPresentation presentation = SCPresentation.Open(pptx);
        IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        IParagraph paragraph = placeholderAutoShape.TextFrame.Paragraphs.Add();

        // Act
        paragraph.Bullet.Type = SCBulletType.Character;
        paragraph.Bullet.Character = "*";
        paragraph.Bullet.Size = 100;
        paragraph.Bullet.FontName = "Tahoma";

        // Assert
        paragraph.Bullet.Type.Should().Be(SCBulletType.Character);
        paragraph.Bullet.Character.Should().Be("*");
        paragraph.Bullet.Size.Should().Be(100);
        paragraph.Bullet.FontName.Should().Be("Tahoma");

        presentation.SaveAs(mStream);

        presentation = SCPresentation.Open(mStream);
        placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        paragraph = placeholderAutoShape.TextFrame.Paragraphs.Last();
        paragraph.Bullet.Type.Should().Be(SCBulletType.Character);
        paragraph.Bullet.Character.Should().Be("*");
        paragraph.Bullet.Size.Should().Be(100);
        paragraph.Bullet.FontName.Should().Be("Tahoma");
    }
        
    [Test]
    public void Type_Setter_sets_Numbered_bullet_type()
    {
        // Arrange
        var mStream = new MemoryStream();
        var pptx = GetInputStream("020.pptx");
        IPresentation presentation = SCPresentation.Open(pptx);
        IAutoShape placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        IParagraph paragraph = placeholderAutoShape.TextFrame.Paragraphs.Add();

        // Act
        paragraph.Bullet.Type = SCBulletType.Numbered;
        paragraph.Bullet.Size = 100;
        paragraph.Bullet.FontName = "Tahoma";

        // Assert
        paragraph.Bullet.Type.Should().Be(SCBulletType.Numbered);
        paragraph.Bullet.Size.Should().Be(100);
        paragraph.Bullet.FontName.Should().Be("Tahoma");

        presentation.SaveAs(mStream);

        presentation = SCPresentation.Open(mStream);
        placeholderAutoShape = (IAutoShape)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        paragraph = placeholderAutoShape.TextFrame.Paragraphs.Last();
        paragraph.Bullet.Type.Should().Be(SCBulletType.Numbered);
        paragraph.Bullet.Size.Should().Be(100);
        paragraph.Bullet.FontName.Should().Be("Tahoma");
    }
}