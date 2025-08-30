using FluentAssertions;
using NUnit.Framework;

namespace ShapeCrawler.DevTests;

public class TableShapeTests
{
    [Test]
    public void Width_Setter_increases_column_widths_proportionally()
    {
        // Arrange
        var pres = new Presentation(p => { p.Slide(s => { s.Table("Table 1", 100, 100, 2, 1); }); });
        var tableShape = pres.Slide(1).Shape("Table 1");
        var newWidth = tableShape.Height * 1.25m;
        var columnWidthBefore = tableShape.Table.Columns.First().Width;

        // Act
        tableShape.Height = newWidth;

        // Assert
        tableShape.Table.Columns.First().Width.Should().BeGreaterThan(columnWidthBefore);
        pres.Validate();
    }
}