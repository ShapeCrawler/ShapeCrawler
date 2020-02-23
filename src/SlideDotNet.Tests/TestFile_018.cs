using System.Linq;
using SlideDotNet.Models;
using Xunit;
// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
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
            var pic = pre.Slides[0].Shapes.Single(x=>x.Id == 7);

            // Act
            var hasPicture = pic.HasPicture;
            var y = pic.Y;
            var picBytes = pic.Picture.ImageEx.GetImageBytes().Result;

            // Assert
            Assert.True(hasPicture);
            Assert.NotNull(picBytes);
            Assert.Equal(4, y);
        }

        [Fact]
        public void Chart_Title_Test()
        {
            // Arrange
            var pre = new Presentation(Properties.Resources._018);
            var chartShape6 = pre.Slides[0].Shapes.Single(x => x.Id == 6);

            // Act
            var title = chartShape6.Chart.Title;

            // Assert
            Assert.Equal("Test title", title);
        }
    }
}
