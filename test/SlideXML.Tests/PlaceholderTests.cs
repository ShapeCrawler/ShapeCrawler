using System.IO;
using NSubstitute;
using SlideXML.Enums;
using SlideXML.Exceptions;
using SlideXML.Models;
using SlideXML.Services.Placeholders;
using Xunit;
// ReSharper disable TooManyDeclarations

namespace SlideXML.Tests
{
    /// <summary>
    /// Represents test for <see cref="PlaceholderData"/> class APIs.
    /// </summary>
    public class PlaceholderTests
    {
        [Fact]
        public void Equals_Test()
        {
            // ARRANGE
            var stubPhXml1 = new PlaceholderXML { PlaceholderType = PlaceholderType.Title };
            var phSl1 = new PlaceholderData(stubPhXml1);
            var stubPhXml2 = new PlaceholderXML { PlaceholderType = PlaceholderType.Title };
            var phSl2 = new PlaceholderData(stubPhXml2);
            var stubPhXml3 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 1};
            var phSl3 = new PlaceholderData(stubPhXml3);
            var phSl4 = new PlaceholderData(stubPhXml3);

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
        public void GetHashCode_Test()
        {
            // ARRANGE
            var stubPhXml1 = new PlaceholderXML { PlaceholderType = PlaceholderType.Title };
            var phSl1 = new PlaceholderData(stubPhXml1);
            var stubPhXml2 = new PlaceholderXML { PlaceholderType = PlaceholderType.Title };
            var phSl2 = new PlaceholderData(stubPhXml2);
            var stubPhXml3 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 1 };
            var stubPhXml4 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 1 };
            var stubPhXml5 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 2 };
            var phSl3 = new PlaceholderData(stubPhXml3);
            var phSl4 = new PlaceholderData(stubPhXml4);
            var phSl5 = new PlaceholderData(stubPhXml5);

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
            var ex = Assert.Throws<PresentationIsLargeException>(() => new Presentation(mockStream));
            var expectedCode = (int)ExceptionCodes.PresentationIsLargeException;
            Assert.Equal(expectedMessage, ex.Message);
            Assert.Equal(expectedCode, ex.ErrorCode);
        }
    }
}
