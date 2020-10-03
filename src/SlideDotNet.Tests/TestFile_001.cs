using System.IO;
using SlideDotNet.Models;
using Xunit;
using System.Linq;
using FluentAssertions;

// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace SlideDotNet.Tests
{
    public class TestFile_001
    {
        [Fact]
        public void Slides_TestSlidesNumber()
        {
            var ms = new MemoryStream(Properties.Resources._001);
            var pre = new PresentationEx(ms);

            // Act
            var sldNumber = pre.Slides.Count();
            pre.Close();
            
            // Assert
            sldNumber.Should().Be(2);
        }
    }
}
