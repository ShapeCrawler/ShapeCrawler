using System.Linq;
using SlideXML.Models;
using Xunit;
// ReSharper disable TooManyChainedReferences

namespace SlideXML.Tests
{
    public class TestFile_015
    {
        [Fact]
        public void FontHeight_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._015);
            var elId5 = pre.Slides[0].Elements.Single(s => s.Id == 5);

            // Act
            var fh = elId5.TextFrame.Paragraphs[0].Portions[2].FontHeight;

            // Arrange
            Assert.Equal(1800, fh);
        }
    }
}
