using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Models;
using SlideDotNet.Models;
using Xunit;


// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.UnitTests
{
    public class TestFile_001
    {
        [Fact]
        public void SlideCustomData_ShouldReturnData_CustomDataIsAssigned()
        {
            // Arrange
            const string customDataString = "Test custom data";
            var origPreStream = new MemoryStream();
            origPreStream.Write(Properties.Resources._001);
            var originPre = new PresentationEx(origPreStream);
            var slide = originPre.Slides.First();

            // Act
            slide.CustomData = customDataString;
            var savedPreStream = new MemoryStream();
            originPre.SaveAs(savedPreStream);
            var savedPre = new PresentationEx(savedPreStream);
            var customData = savedPre.Slides.First().CustomData;

            // Assert
            customData.Should().Be(customDataString);
        }
    }
}
