using Fixture;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class TableShapeTests : SCTest
{
    private readonly Fixtures fixtures = new();

    [Test]
    public void Width_Setter_increases_column_widths_proportionally()
    {
        // Arrange
        var shapeName = Fixtures.String();
        var pres = new Presentation(p =>
        {
            p.Slide(s =>
            {
                s.Table(shapeName,fixtures.Int(), fixtures.Int(), fixtures.Int(), fixtures.Int());
            });
        });
        var tableShape = pres.Slide(1).Shape(shapeName);
        var newShapeWidth = tableShape.Width + fixtures.Int();
        var column = tableShape.Table.Columns.First();
        var columnWidthBefore = column.Width;

        // Act
        tableShape.Width = newShapeWidth;

        // Assert
        column.Width.Should().BeGreaterThan(columnWidthBefore);
        ValidatePresentation(pres);
    }
}