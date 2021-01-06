using System.Linq;
using ShapeCrawler.Models;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.Tests.Unit
{
    public class TestFile_020
    {
        [Fact]
        public void FontHeight_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._020);

            // Act
            var shape3 = pre.Slides[0].Shapes.Single(x => x.Id == 3);

            var text = shape3.TextFrame.Text;
            var fh = shape3.TextFrame.Paragraphs.First().Portions.First().Font.Size;

            // Assert
            Assert.Equal(1800, fh);
        }
    }
}
