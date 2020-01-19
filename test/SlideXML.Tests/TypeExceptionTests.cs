using SlideXML.Exceptions;
using Xunit;

namespace SlideXML.Tests
{
    /// <summary>
    /// Contains unit tests for the <see cref="TypeException"/> class.
    /// </summary>
    public class TypeExceptionTests
    {
        [Fact]
        public void Constructor_Test()
        {
            // ACT
            var exception = Assert.ThrowsAsync<TypeException>(() => throw new TypeException()).Result;

            // ASSERT
            Assert.Equal(101, exception.ErrorCode);
        }
    }
}
