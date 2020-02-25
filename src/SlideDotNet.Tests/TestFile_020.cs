using System.Linq;
using SlideDotNet.Models;
using Xunit;
// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_020
    {
        [Fact]
        public void FontHeight_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._020);

            // Act
            var shape3 = pre.Slides[0].Shapes.Single(x => x.Id == 3);

            var text = shape3.TextFrame.Text;
            var fh = shape3.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            // Assert
            Assert.Equal(1800, fh);
        }
    }
}
