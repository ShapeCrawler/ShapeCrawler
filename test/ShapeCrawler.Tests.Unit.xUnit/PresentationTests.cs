using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Charts;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class PresentationTests : SCTest
{
    [Xunit.Theory]
    [MemberData(nameof(TestCasesSlidesRemove))]
    public void Slides_Remove_removes_slide(string file, int expectedSlidesCount)
    {
        // Arrange
        var pptx = StreamOf(file);
        var pres = new Presentation(pptx);
        var removingSlide = pres.Slides[0];
        var mStream = new MemoryStream();

        // Act
        pres.Slides.Remove(removingSlide);

        // Assert
        pres.Slides.Should().HaveCount(expectedSlidesCount);

        pres.SaveAs(mStream);
        pres = new Presentation(mStream);
        pres.Slides.Should().HaveCount(expectedSlidesCount);
    }
    
    public static IEnumerable<object[]> TestCasesSlidesRemove()
    {
        yield return new object[] {"007_2 slides.pptx", 1};
        yield return new object[] {"006_1 slides.pptx", 0};
    }
}