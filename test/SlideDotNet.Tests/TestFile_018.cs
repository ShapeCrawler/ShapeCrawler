using System.Linq;
using SlideDotNet.Models;
using Xunit;
// ReSharper disable TooManyDeclarations

// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_018
    {
        [Fact]
        public void Picture_Placeholder_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._018);
            var pic = pre.Slides[0].Elements.Single(x=>x.Id == 7);

            // Act
            var hasPicture = pic.HasPicture;
            var y = pic.Y;
            var picBytes = pic.Picture.ImageEx.GetBytes();

            // Assert
            Assert.True(hasPicture);
            Assert.NotNull(picBytes);
            Assert.Equal(4, y);
        }
    }
}
