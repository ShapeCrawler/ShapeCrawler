using ShapeCrawler.Models;
using SlideDotNet.Models;
using Xunit;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace ShapeCrawler.UnitTests
{
    public class TestFile_023
    {
        [Fact]
        public void Slide_Shapes_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._023);

            // Act-Assert
            var shapes = pre.Slides[0].Shapes;
        }
    }
}
