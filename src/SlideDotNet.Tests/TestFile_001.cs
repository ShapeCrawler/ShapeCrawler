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
        public void SlidesCount_ShouldReturnTwo_PresentationContainsTwoSlides()
        {
            // Arrange
            var pre = new PresentationEx(Properties.Resources._001);

            // Act
            var sldNumber = pre.Slides.Count();
            
            // Assert
            sldNumber.Should().Be(2);
        }

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

        [Fact]
        public void SlideCustomData_ShouldReturnNull_CustomDataIsNotAssigned()
        {
            // Arrange
            var preStream = new MemoryStream();
            preStream.Write(Properties.Resources._001, 0, Properties.Resources._001.Length);
            var pre = new PresentationEx(preStream);
            var slide = pre.Slides.First();

            // Act
            var customData = slide.CustomData;

            // Assert
            customData.Should().BeNull();
        }
    }
}
