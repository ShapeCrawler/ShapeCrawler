using SlideXML.Enums;
using SlideXML.Services.Placeholders;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Represents test for <see cref="PlaceholderSL"/> class APIs.
    /// </summary>
    public class PlaceholderSLTests
    {
        [Fact]
        public void Equals_Test()
        {
            // ARRANGE
            var stubPhXml1 = new PlaceholderXML { PlaceholderType = PlaceholderType.Title };
            var phSl1 = new PlaceholderSL(stubPhXml1);
            var stubPhXml2 = new PlaceholderXML { PlaceholderType = PlaceholderType.Title };
            var phSl2 = new PlaceholderSL(stubPhXml2);
            var stubPhXml3 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 1};
            var phSl3 = new PlaceholderSL(stubPhXml3);
            var phSl4 = new PlaceholderSL(stubPhXml3);

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
            var phSl1 = new PlaceholderSL(stubPhXml1);
            var stubPhXml2 = new PlaceholderXML { PlaceholderType = PlaceholderType.Title };
            var phSl2 = new PlaceholderSL(stubPhXml2);
            var stubPhXml3 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 1 };
            var stubPhXml4 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 1 };
            var stubPhXml5 = new PlaceholderXML { PlaceholderType = PlaceholderType.Custom, Index = 2 };
            var phSl3 = new PlaceholderSL(stubPhXml3);
            var phSl4 = new PlaceholderSL(stubPhXml4);
            var phSl5 = new PlaceholderSL(stubPhXml5);

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
    }
}
