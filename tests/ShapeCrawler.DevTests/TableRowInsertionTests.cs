using System.IO;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class TableRowInsertionTests : SCTest
{
    [Test]
    public void Rows_Add_adds_row_at_the_specified_index()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Table("Table 1");
        var rowsCountBefore = table.Rows.Count;

        // Act
        table.Rows.Add(1);

        // Assert
        table.Rows.Should().HaveCount(rowsCountBefore + 1);
        table.Rows[1].Cells[0].TextBox.Text.Should().BeEmpty();
        pres = SaveAndOpenPresentation(pres);
        table = pres.Slide(1).Table("Table 1");
        table.Rows.Should().HaveCount(rowsCountBefore + 1);
        pres.Validate();
    }

    [Test]
    public void Rows_Add_adds_a_new_row_at_the_specified_index_with_the_template_height()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Table("Table 1");
        var templateRowIndex = 0;
        var templateRowHeight = table.Rows[templateRowIndex].Height;

        // Act
        table.Rows.Add(1, templateRowIndex);

        // Assert
        pres = SaveAndOpenPresentation(pres);
        table = pres.Slide(1).Table("Table 1");
        table.Rows[1].Height.Should().Be(templateRowHeight);
        pres.Validate();
    }
}
