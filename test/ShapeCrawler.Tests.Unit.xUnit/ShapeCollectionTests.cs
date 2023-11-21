using System.Diagnostics.CodeAnalysis;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;
using Assert = Xunit.Assert;

// ReSharper disable SuggestVarOrType_BuiltInTypes
// ReSharper disable TooManyChainedReferences
// ReSharper disable TooManyDeclarations

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class ShapeCollectionTests : SCTest
{
    [Xunit.Theory]
    [LayoutShapeData("autoshape-case004_subtitle.pptx", slideNumber: 1, shapeName: "Group 1")]
    [MasterShapeData("autoshape-case004_subtitle.pptx", shapeName: "Group 1")]
    public void GetByName_returns_shape_by_specified_name(IShape shape)
    {
        // Arrange
        var groupShape = (IGroupShape)shape;
        var shapeCollection = groupShape.Shapes;
            
        // Act
        var resultShape = shapeCollection.GetByName<IShape>("AutoShape 1");

        // Assert
        resultShape.Should().NotBeNull();
    }

    [Xunit.Theory]
    [SlideData("#1", "002.pptx", slideNumber: 1, expectedResult: 4)]
    [SlideData("#2","003.pptx", slideNumber: 1, expectedResult: 5)]
    [SlideData("#3","013.pptx", slideNumber: 1, expectedResult: 4)]
    [SlideData("#4","023.pptx", slideNumber: 1, expectedResult: 1)]
    [SlideData("#5","014.pptx", slideNumber: 3, expectedResult: 5)]
    [SlideData("#6","009_table.pptx", slideNumber: 1, expectedResult: 6)]
    [SlideData("#7","009_table.pptx", slideNumber: 2, expectedResult: 8)]
    public void Count_returns_number_of_shapes(string label, ISlide slide, int expectedCount)
    {
        // Arrange
        var shapeCollection = slide.Shapes;
            
        // Act
        int shapesCount = shapeCollection.Count;

        // Assert
        shapesCount.Should().Be(expectedCount);
    }
}