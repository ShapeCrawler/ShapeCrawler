﻿using System;
using System.Linq;
using SlideDotNet.Enums;
using SlideDotNet.Models;
using Xunit;
// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_019Fixture : IDisposable
    {
        public PresentationEx pre019 { get; }

        public TestFile_019Fixture()
        {
            pre019 = new PresentationEx(Properties.Resources._019);
        }

        public void Dispose()
        {
            pre019.Close();
        }
    }

    public class TestFile_019 : IClassFixture<TestFile_019Fixture>
    {
        private readonly TestFile_019Fixture _fixture;

        public TestFile_019(TestFile_019Fixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void AutoShape_FontHeight()
        {
            // Arrange
            var pre = _fixture.pre019;

            // Act
            var fh = pre.Slides[0].Shapes.Single(x=>x.Id == 4103).TextFrame.Paragraphs.First().Portions.First().FontHeight;

            // Assert
            Assert.Equal(1800, fh);
        }

        [Fact]
        public void Chart_Title_Test()
        {
            // Arrange
            var pre = _fixture.pre019;

            // Act
            var chartTitle = pre.Slides[0].Shapes.Single(x => x.Id == 4).Chart.Title;

            // Assert
            Assert.Equal("Test title", chartTitle);
        }

        [Fact]
        public void Picture_DoNotParseStrangePicture_Test()
        {
            // Arrange
            var pre = _fixture.pre019;

            // Act - Assert
            Assert.ThrowsAny<Exception>(() => pre.Slides[1].Shapes.Single(x => x.Id == 47));
        }

        [Fact]
        public void SlideNumber_Test()
        {
            // Arrange
            var pre = _fixture.pre019;
            var shape2 = pre.Slides[0].Shapes.Single(x => x.Id == 2);

            // Act
            var text = shape2.TextFrame.Text;
            var phType = shape2.PlaceholderType;
            var fh = shape2.TextFrame.Paragraphs.First().Portions.First().FontHeight;

            // Arrange
            Assert.Equal("1", text);
            Assert.Equal(PlaceholderType.SlideNumber, phType);
            Assert.Equal(1200, fh);
        }
    }
}
