using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Tables;
using System.Linq;

namespace ShapeCrawler.DevTests;

public class Issue1185Tests : SCTest
{
    [Test]
    public void Rows_Add_should_copy_style_from_template_row()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;
        var templateRow = table.Rows[0];
        var templateCell = templateRow.Cells[0];
        var expectedColor = "FF0000";
        templateCell.Fill.SetColor(expectedColor);
        
        // Act
        table.Rows.Add(1, 0); // Add new row at index 1, using row 0 as template
        pres.Save(@"c:\temp\after.pptx");
        
        // Assert
        var newRow = table.Rows[1];
        var newCell = newRow.Cells[0];
        newCell.Fill.Color.Should().Be(expectedColor);
    }
}
