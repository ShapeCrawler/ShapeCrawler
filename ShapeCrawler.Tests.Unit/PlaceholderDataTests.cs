using ShapeCrawler.Placeholders;
using Xunit;

// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit
{
    /// <summary>
    /// Represents test for <see cref="PlaceholderLocationData"/> class APIs.
    /// </summary>
    public class PlaceholderDataTests
    {
        [Fact]
        public void Equals_ReturnsFalse_WhenPlaceholdersAreNotSame()
        {
            // Arrange
            var phData1 = new PlaceholderData { Index = 4, PlaceholderType = PlaceholderType.Custom };
            var phData2 = new PlaceholderData { Index = 4, PlaceholderType = PlaceholderType.SlideNumber };
            var phData3 = new PlaceholderData { PlaceholderType = PlaceholderType.Title };
            var phData4 = new PlaceholderData { PlaceholderType = PlaceholderType.Custom, Index = 1 };
            var phData5 = new PlaceholderData { PlaceholderType = PlaceholderType.Custom, Index = 2 };
            var phLocationData1 = new PlaceholderLocationData(phData3);
            var phLocationData2 = new PlaceholderLocationData(phData4);
            var phLocationData3 = new PlaceholderLocationData(phData5);

            // Act
            var isEqualsCase1 = phLocationData1.Equals(phLocationData2);
            var isEqualsCase2 = phData1.Equals(phData2);
            var isEqualsCase3 = phLocationData2.Equals(phLocationData3);

            // Assert
            Assert.False(isEqualsCase1);
            Assert.False(isEqualsCase2);
            Assert.False(isEqualsCase3);
        }
    }
}
