using System.Linq;
using SlideDotNet.Models;
using Xunit;
// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_021
    {
        [Fact]
        public void Shape_Fill_DoNotThrowException_Test()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._021);
            var sp108 = pre.Slides[0].Shapes.Single(x => x.Id == 108);

            // Act-Assert
            var fill = sp108.Fill;
        }
    }
}
