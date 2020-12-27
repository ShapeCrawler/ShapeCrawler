using System.IO;
using NSubstitute;
using ShapeCrawler.Enums;
using ShapeCrawler.Exceptions;
using ShapeCrawler.Models;
using ShapeCrawler.Services.Placeholders;
using ShapeCrawler.Statics;
using Xunit;

// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    /// <summary>
    /// Represents test for <see cref="PlaceholderLocationData"/> class APIs.
    /// </summary>
    public class PlaceholderTests
    {
        [Fact]
        public void Equals_Test_Case1()
        {
            // ARRANGE
            var stubPhXml1 = new PlaceholderData { PlaceholderType = PlaceholderType.Title };
            var phSl1 = new PlaceholderLocationData(stubPhXml1);
            var stubPhXml2 = new PlaceholderData { PlaceholderType = PlaceholderType.Title };
            var phSl2 = new PlaceholderLocationData(stubPhXml2);
            var stubPhXml3 = new PlaceholderData { PlaceholderType = PlaceholderType.Custom, Index = 1};
            var phSl3 = new PlaceholderLocationData(stubPhXml3);
            var phSl4 = new PlaceholderLocationData(stubPhXml3);

            // ACT
            var isEquals1 = phSl1.Equals(phSl2);
            var isEquals2 = phSl1.Equals(phSl3);
            var isEquals3 = phSl3.Equals(phSl4);

            // ASSERT
            Assert.True(isEquals1);
            Assert.False(isEquals2);
            Assert.True(isEquals3);
        }

        [Fact]
        public void Equals_Test_Case2()
        {
            // ARRANGE
            var customIndex4 = new PlaceholderData
            {
                Index = 4,
                PlaceholderType = PlaceholderType.Custom
            };
            var sldNumIndex4 = new PlaceholderData
            {
                Index = 4,
                PlaceholderType = PlaceholderType.SlideNumber
            };

            // ACT
            var isEquals4 = customIndex4.Equals(sldNumIndex4);

            // ASSERT
            Assert.False(isEquals4);
        }

        [Fact]
        public void GetHashCode_Test()
        {
            // ARRANGE
            var stubPhXml1 = new PlaceholderData { PlaceholderType = PlaceholderType.Title };
            var phSl1 = new PlaceholderLocationData(stubPhXml1);
            var stubPhXml2 = new PlaceholderData { PlaceholderType = PlaceholderType.Title };
            var phSl2 = new PlaceholderLocationData(stubPhXml2);
            var stubPhXml3 = new PlaceholderData { PlaceholderType = PlaceholderType.Custom, Index = 1 };
            var stubPhXml4 = new PlaceholderData { PlaceholderType = PlaceholderType.Custom, Index = 1 };
            var stubPhXml5 = new PlaceholderData { PlaceholderType = PlaceholderType.Custom, Index = 2 };
            var phSl3 = new PlaceholderLocationData(stubPhXml3);
            var phSl4 = new PlaceholderLocationData(stubPhXml4);
            var phSl5 = new PlaceholderLocationData(stubPhXml5);

            // ACT
            var hash1 = phSl1.GetHashCode();
            var hash2 = phSl2.GetHashCode();
            var hash3 = phSl3.GetHashCode();
            var hash4 = phSl4.GetHashCode();
            var hash5 = phSl5.GetHashCode();

            // ASSERT
            Assert.Equal(hash1, hash2);
            Assert.Equal(hash3, hash4);
            Assert.NotEqual(hash3, hash5);
        }

        [Fact]
        public void Constructor_Test()
        {
            // Arrange
            var mockStream = Substitute.For<Stream>();
            var maxLength = Limitations.MaxPresentationSize;
            var stubStreamLength = Limitations.MaxPresentationSize + 1;
            mockStream.Length.Returns(stubStreamLength);
            var expectedMessage = $"The size of presentation more than {maxLength} bytes.";

            // Act-Assert
            var ex = Assert.Throws<PresentationIsLargeException>(() => new Presentation(mockStream, false));
            var expectedCode = (int)ExceptionCodes.PresentationIsLargeException;
            Assert.Equal(expectedMessage, ex.Message);
            Assert.Equal(expectedCode, ex.ErrorCode);
        }
    }
}
