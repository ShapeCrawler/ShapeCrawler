﻿using System;
using System.Linq;
using SlideDotNet.Models;
using Xunit;
// ReSharper disable TooManyDeclarations
// ReSharper disable InconsistentNaming
// ReSharper disable TooManyChainedReferences

namespace SlideDotNet.Tests
{
    public class TestFile_018Fixture : IDisposable
    {
        public PresentationEx pre018 { get; }

        public TestFile_018Fixture()
        {
            pre018 = new PresentationEx(Properties.Resources._018);
        }

        public void Dispose()
        {
            pre018.Close();
        }
    }

    public class TestFile_018 : IClassFixture<TestFile_018Fixture>
    {
        private readonly TestFile_018Fixture _fixture;

        public TestFile_018(TestFile_018Fixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void Picture_Placeholder_Test()
        {
            // Arrange
            var pre = _fixture.pre018;
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
            var pre = _fixture.pre018;
            var chartShape6 = pre.Slides[0].Shapes.Single(x => x.Id == 6);

            // Act
            var title = chartShape6.Chart.Title;

            // Assert
            Assert.Equal("Test title", title);
        }
    }
}
