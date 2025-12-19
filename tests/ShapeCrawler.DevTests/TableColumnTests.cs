using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;

namespace ShapeCrawler.DevTests;

public class TableColumnTests : SCTest
{
    [Test]
    public void Columns_RemoveAt_removes_column_by_specified_index()
    {
        // Arrange
        var ms = new MemoryStream();
        var pptx = TestAsset("table-case001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(1).Shape("Table 1").Table;
        var expectedColumnsCount = table.Columns.Count - 1;

        // Act
        table.Columns.RemoveAt(1);

        // Assert
        table.Columns.Should().HaveCount(expectedColumnsCount);
        pres.Save(ms);
        pres = new Presentation(ms);
        table = pres.Slide(1).Shape("Table 1").Table;
        table.Columns.Should().HaveCount(expectedColumnsCount);
        ValidatePresentation(pres);
    }

    [Test]
    public void Columns_Add_adds_column()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;
        var expectedColumnsCount = table.Columns.Count + 1;

        // Act
        table.Columns.Add();

        // Assert
        table.Columns.Should().HaveCount(expectedColumnsCount);
        ValidatePresentation(pres);
    }

    [Test]
    public void Columns_InsertAfter_inserts_column_after_the_specified_column_number()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;

        // Act
        table.Columns.InsertAfter(1);
        var cell = table.Cell(1, 2);

        // Assert
        cell.ShapeText.Text.Should().BeEmpty("because before adding column the cell (1,2) was not empty.");
        ValidatePresentation(pres);
    }

    [Test]
    public void Columns_Add_sets_width_of_all_columns_proportionally()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case003.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;
        var columnsCountBefore = table.Columns.Count;
        var columnWidthBefore = table.Columns.Select(c => c.Width).ToList();
        var totalWidthBefore = table.Columns.Sum(c => c.Width);
        var newTotalWidth = totalWidthBefore + table.Columns[columnsCountBefore - 1].Width;

        // Act
        table.Columns.Add();

        // Assert
        var widthRatio = totalWidthBefore / newTotalWidth;
        table.Columns.Select(column => (int)column.Width).ToList().Take(columnsCountBefore).Should()
            .BeEquivalentTo(columnWidthBefore.Select(w => (int)(w * widthRatio)));
    }

    [Test]
    public void Columns_InsertAfter_sets_width_of_all_columns_proportionally()
    {
        // Arrange
        var pptx = TestAsset("table-case003.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shape("Table 1").Table;
        var columnsCountBefore = table.Columns.Count;
        var columnWidthBefore = table.Columns.Select(c => c.Width).ToList();
        var totalWidthBefore = table.Columns.Sum(c => c.Width);
        var newTotalWidth = totalWidthBefore + table.Columns[2].Width;

        // Act
        table.Columns.InsertAfter(3);

        // Assert
        var widthRatio = totalWidthBefore / newTotalWidth;
        table.Columns.Select(c => (int)c.Width).ToList().Take(columnsCountBefore).Should()
            .BeEquivalentTo(columnWidthBefore.Select(w => (int)(w * widthRatio)));
    }

    [Test]
    public void Columns_Duplicate_increases_column_count_by_one()
    {
        // Arrange
        var pptx = TestAsset("table-case001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(1).Shape("Table 1").Table;
        var column = table.Columns[0];
        var columnsCountBefore = table.Columns.Count;

        // Act
        column.Duplicate();

        // Assert
        table.Columns.Should().HaveCount(columnsCountBefore + 1);
        ValidatePresentation(pres);
    }

    [Test]
    public void Columns_Duplicate_copies_column_with_all_its_cells()
    {
        // Arrange
        var pptx = TestAsset("table-case002.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(1).Shape("Table 1").Table;
        var column = table.Columns[0];

        // Act
        column.Duplicate();

        // Assert
        foreach (var row in table.Rows)
        {
            row.Cells.Should().HaveCount(table.Columns.Count);
            row.Cells[0].ShapeText.Text.Should().Be(row.Cells[table.Columns.Count - 1].ShapeText.Text);
        }
    }

    [Test]
    public void Columns_Duplicate_copies_middle_column_with_all_its_cells()
    {
        // Arrange
        var pptx = TestAsset("table-case003.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(1).Shape("Table 1").Table;
        var column = table.Columns[1];

        // Act
        column.Duplicate();

        // Assert
        foreach (var row in table.Rows)
        {
            row.Cells.Should().HaveCount(table.Columns.Count);
            row.Cells[1].ShapeText.Text.Should().Be(row.Cells[table.Columns.Count - 1].ShapeText.Text);
        }
    }


    [Test]
    public void Columns_Duplicate_sets_width_of_all_columns_proportionally()
    {
        // Arrange
        var pptx = TestAsset("table-case003.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(1).Shape("Table 1").Table;
        var column = table.Columns[0];
        var columWidthBefore = table.Columns.Select(c => c.Width).ToList();
        var totalWidthBefore = table.Columns.Sum(c => c.Width);
        var newTotalWidth = totalWidthBefore + column.Width;

        // Act
        column.Duplicate();

        // Assert
        var widthRatio = totalWidthBefore / newTotalWidth;
        table.Columns.Select(c => (int)c.Width).ToList().Take(columWidthBefore.Count).Should()
            .BeEquivalentTo(columWidthBefore.Select(w => (int)(w * widthRatio)));
    }

    [Test]
    public void ColumnsCount_ReturnsNumberOfColumnsInTheTable()
    {
        // Arrange
        var table = new Presentation(TestAsset("001.pptx")).Slide(2).Shape(4).Table;

        // Act
        int columnsCount = table.Columns.Count;

        // Assert
        columnsCount.Should().Be(3);
    }

    [Test]
    public void Column_Width_Getter()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slide(2).Shape(4).Table;

        // Act & Assert
        table.Columns[0].Width.Should().BeApproximately(275.99m, 0.01m);
    }

    [Test]
    public void Column_Width_Setter_sets_width_of_column()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slide(2).Shape(3).Table;
        const int newColumnWidth = 427;
        var mStream = new MemoryStream();

        // Act
        table.Columns[0].Width = newColumnWidth;

        // Assert
        table.Columns[0].Width.Should().Be(newColumnWidth);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = pres.Slide(2).Shape(3).Table;
        table.Columns[0].Width.Should().Be(newColumnWidth);
    }
}