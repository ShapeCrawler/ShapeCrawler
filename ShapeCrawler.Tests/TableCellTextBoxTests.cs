using FluentAssertions;
using ShapeCrawler.Tests.Helpers;

namespace ShapeCrawler.Tests;

public class TableCellTextBoxTests : SCTest
{
    [TestCase("Alice")]
    [TestCase("")]
    public void ReplaceText_updates_table_cells_and_survives_reopening(string value)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Shapes.AddTable(0, 0, 1, 1);
        var cell = pres.Slides[0].Shapes.Last().Table.Rows[0].Cells[0];
        cell.TextBox.SetText("Hello {{name}}!");

        // Act
        var count = cell.TextBox.ReplaceText("{{name}}", value);

        // Assert
        count.Should().Be(1);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        reopened.Slides[0].Shapes.Last().Table.Rows[0].Cells[0].TextBox.Text.Should().Be($"Hello {value}!");
        ValidatePresentation(reopened);
    }
}
