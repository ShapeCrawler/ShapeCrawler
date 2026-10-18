using FluentAssertions;
using System.Text.RegularExpressions;
using ShapeCrawler.Tests.Helpers;

namespace ShapeCrawler.Tests;

public class TableCellTextBoxTests : SCTest
{
    [TestCase(false)]
    [TestCase(true)]
    public void HighlightText_and_HighlightRegex_round_trip_cell_matches(bool regex)
    {
        // Arrange
        using var pres = new Presentation(p => p.Slide());
        pres.Slides[0].Shapes.AddTable(0, 0, 1, 1);
        var box = pres.Slides[0].Shapes.Last().Table[0, 0].TextBox;
        box.SetText("a TAG b TAG c");
        box.Paragraphs[0].Portions[0].Font.IsBold = true;

        // Act
        var count = regex ? box.HighlightRegex(new Regex(@"\bTAG\b"), new Color("FFFF00"))
            : box.HighlightText("TAG", new Color("FFFF00"));

        // Assert
        count.Should().Be(2);
        using var saved = new MemoryStream();
        pres.Save(saved);
        using var reopened = new Presentation(saved);
        var result = reopened.Slides[0].Shapes.Last().Table[0, 0].TextBox;
        result.Text.Should().Be("a TAG b TAG c");
        result.Paragraphs[0].Portions.Where(portion => !portion.TextHighlightColor.IsTransparent)
            .Select(portion => portion.Text).Should().Equal("TAG", "TAG");
        result.Paragraphs[0].Portions.Should().OnlyContain(portion => portion.Font.IsBold);
        ValidatePresentation(reopened);
    }

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
