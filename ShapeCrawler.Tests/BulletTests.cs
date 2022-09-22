#if DEBUG

using System.IO;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests
{
    public class BulletTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public BulletTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void ChangeType()
        {
            // Arrange
            var presentationStream = GetTestStream("autoshape-case003.pptx");
            using var presentation = SCPresentation.Open(presentationStream);
            var shape = presentation.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
            var bullet = shape.TextFrame!.Paragraphs[0].Bullet;

            // Act
            bullet.Type = SCBulletType.Character;
            bullet.Character = "*";

            // Assert
            bullet.Type.Should().Be(SCBulletType.Character);
            bullet.Character.Should().Be("*");

            var savedPreStream = new MemoryStream();
            presentation.SaveAs(savedPreStream);
            var newPresentation = SCPresentation.Open(savedPreStream);
            shape = newPresentation.Slides[0].Shapes.GetByName<IAutoShape>("AutoShape 1");
            bullet = shape.TextFrame!.Paragraphs[0].Bullet;
            bullet.Type.Should().Be(SCBulletType.Character);
            bullet.Character.Should().Be("*");

        }

    }
}

#endif