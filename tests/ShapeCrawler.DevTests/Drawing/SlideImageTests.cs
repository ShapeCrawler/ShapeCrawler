using System.IO;
using FluentAssertions;
using ShapeCrawler;
using NUnit.Framework;

using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests.Drawing;

public class SlideImageTests : SCTest
{
    [Test]
    public void SaveAsPng_GeneratesValidImage()
    {
        // Arrange
        var pres = new Presentation();
        pres.Slides.Add(1);
        var slide = pres.Slides[0];
        using var stream = new MemoryStream();

        // Act
        slide.SaveAsPng(stream);

        // Assert
        stream.Length.Should().BeGreaterThan(0);
    }
}
