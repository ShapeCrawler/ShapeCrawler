using FluentAssertions;
using ShapeCrawler.Tests.Shared;
using Xunit;

namespace ShapeCrawler.UnitTests;

public class AutoShapeTests
{
#if DEBUG
    [Fact(Skip = "On Hold (https://github.com/ShapeCrawler/ShapeCrawler/issues/444)")]
    public void Duplicate_duplicates_AutoShape()
    {
        // Arrange
        var pptx = TestHelper.GetStream("autoshape-grouping.pptx");
        var pres = SCPresentation.Open(pptx);
        var shape = pres.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 6");

        // Act
        var shapeCopy = shape.Duplicate();

        // Assert
        shapeCopy.X.Should().Be(12);
        shapeCopy.Width.Should().Be(shape.Width);
        shapeCopy.TextFrame.Text.Should().Be(shapeCopy.TextFrame.Text);
    }
#endif
}