using System.Linq;
using SlideDotNet.Models;
using Xunit;

// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_015
    {
        [Fact]
        public void FontHeight_Test_1()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._015);
            var elId5 = pre.Slides[0].Elements.Single(s => s.Id == 5);

            // Act
            var fh = elId5.TextFrame.Paragraphs[0].Portions[2].FontHeight;

            // Arrange
            Assert.Equal(1800, fh);
        }

        [Fact]
        public void Placeholder_FontHeight_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._015);
            var el = pre.Slides[1].Elements.Single(s => s.Id == 61);

            // Act
            var fh = el.TextFrame.Paragraphs[0].Portions[0].FontHeight;

            // Assert
            Assert.Equal(1867, fh);
        }
    }
}
