using System.Linq;
using SlideDotNet.Enums;
using SlideDotNet.Models;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_023
    {
        [Fact]
        public void Slide_Shapes_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._023);

            // Act-Assert
            var shapes = pre.Slides[0].Shapes;
        }
    }
}
