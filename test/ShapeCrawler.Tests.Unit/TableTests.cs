using System.Diagnostics.CodeAnalysis;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tests.Unit.Helpers;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tests.Unit;

[SuppressMessage("Usage", "xUnit1013:Public method should be marked as test")]
public class TableTests : SCTest
{
    [Test]
    public void RemoveColumnAt_removes_column_by_specified_index()
    {
        // Arrange
        var ms = new MemoryStream();
        var pptx = GetInputStream("table-case001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var expectedColumnsCount = table.Columns.Count - 1;

        // Act
        table.RemoveColumnAt(1);

        // Assert
        table.Columns.Should().HaveCount(expectedColumnsCount);
        pres.SaveAs(ms);
        pres = SCPresentation.Open(ms);
        table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        table.Columns.Should().HaveCount(expectedColumnsCount);
        PptxValidator.Validate(pres);
    }

    [Test]
    public void Rows_RemoveAt_removes_row_with_specified_index()
    {
        // Arrange
        var pptx = GetInputStream("009_table.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = (ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 3);
        int originRowsCount = table.Rows.Count;
        var mStream = new MemoryStream();

        // Act
        table.Rows.RemoveAt(0);

        // Assert
        table.Rows.Should().HaveCountLessThan(originRowsCount);
        pres.SaveAs(mStream);
        table = (ITable)SCPresentation.Open(mStream).Slides[2].Shapes.First(sp => sp.Id == 3);
        table.Rows.Should().HaveCountLessThan(originRowsCount);
    }

    [Test]
    public void Rows_Add_adds_row()
    {
        // Arrange
        var pptx = GetInputStream("table-case001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");

        // Act
        table.Rows.Add();

        // Assert
        table.Rows.Should().HaveCount(2);
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }

    [Test]
    public void Row_Cells_Count_returns_number_of_cells_in_the_row()
    {
        // Arrange
        var pptx = GetInputStream("009_table.pptx");
        var table = (ITable)SCPresentation.Open(pptx).Slides[2].Shapes
            .First(sp => sp.Id == 3);

        // Act
        var cellsCount = table.Rows[0].Cells.Count;

        // Assert
        cellsCount.Should().Be(3);
    }

    [Test]
    public void Row_Height_Getter_returns_row_height_in_points()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 3);

        // Act
        var rowHeight = table.Rows[0].Height;

        // Act-Assert
        rowHeight.Should().Be(29);
    }

    [Test]
    public void ColumnsCount_ReturnsNumberOfColumnsInTheTable()
    {
        // Arrange
        ITable table = (ITable)SCPresentation.Open(GetInputStream("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);

        // Act
        int columnsCount = table.Columns.Count;

        // Assert
        columnsCount.Should().Be(3);
    }

    [Test]
    public void Column_Width_Getter_returns_width_of_column_in_pixels()
    {
        // Arrange
        var table = (ITable)SCPresentation.Open(GetInputStream("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);

        // Act
        var columnWidth = table.Columns[0].Width;

        // Assert
        columnWidth.Should().Be(367);
    }

    [Test]
    public void Column_Width_Setter_sets_width_of_column()
    {
        // Arrange
        var pres = SCPresentation.Open(GetInputStream("001.pptx"));
        var table = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 3);
        const int newColumnWidth = 427;
        var mStream = new MemoryStream();

        // Act
        table.Columns[0].Width = newColumnWidth;

        // Assert
        table.Columns[0].Width.Should().Be(newColumnWidth);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        table = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 3);
        table.Columns[0].Width.Should().Be(newColumnWidth);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_True_When_cell_is_merged()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var row = pres.Slides[1].Shapes.GetByName<ITable>("Table 4").Rows[1];
        var cell1X0 = row.Cells[0];
        var cell1X1 = row.Cells[1];

        // Act
        var isMerged1 = cell1X0.IsMergedCell;
        var isMerged2 = cell1X1.IsMergedCell;

        // Act-Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        cell1X0.Should().BeSameAs(cell1X1);
    }

    [Test]
    public void Row_Clone_cloning_row_increases_row_count_by_one()
    {
        // Arrange
        var pptx = GetInputStream("tables-case001.pptx");
        var pres = SCPresentation.Open(pptx);
        var targetTable = pres.Slides.First().Shapes.OfType<ITable>().FirstOrDefault();
        var rowCountBefore = targetTable.Rows.Count;
        var row = targetTable.Rows.Last();

        // Act
        row.Clone();

        // Assert
        var rowCountAfter = targetTable.Rows.Count;
        rowCountAfter.Should().Be(rowCountBefore + 1);
    }

    [Test]
    public void Row_Height_Setter_sets_height_of_table_row_in_points()
    {
        // Arrange
        var pptx = GetInputStream("table-case001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var row = table.Rows[0];

        // Act
        row.Height = 58;

        // Assert
        row.Height.Should().Be(58);
        table.Height.Should().Be(76, "because table height was 38px.");
    }

    [Test]
    public void MergeCells_Merges0x0And0x1CellsOf2x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].TextFrame.Text.Should().Be($"id5{Environment.NewLine}Text0_1");

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].TextFrame.Text.Should().Be($"id5{Environment.NewLine}Text0_1");
    }

    [Test]
    public void MergeCells_Merges0x1And0x2CellsOf3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 1], table[0, 2]);

        // Assert
        AssertTable(table);
        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        AssertTable(table);

        static void AssertTable(ITable tableSc)
        {
            tableSc[0, 1].IsMergedCell.Should().BeTrue();
            tableSc[0, 2].IsMergedCell.Should().BeTrue();
            tableSc[0, 1].TextFrame.Text.Should().Be("Text0_2");
            tableSc[0, 2].TextFrame.Text.Should().Be("Text0_2");
        }
    }

    [Test]
    public void MergeCells_Merges0x0And0x1And0x2CellsOf3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 2]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();
    }

    [Test]
    public void MergeCells_Merges0x0And0x1MergedCellsWith0x2CellIn3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 2]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();
    }

    [Test]
    public void MergeCells_merges_0x0_and_1x0_cells_of_2x2_table()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[2].Shapes.GetById<ITable>(5);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[1, 0]);

        // Assert
        AssertTable(table);
        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        table = (ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 5);
        AssertTable(table);

        void AssertTable(ITable table)
        {
            string expectedText = $"id5{Environment.NewLine}Text1_0";
            table[0, 0].IsMergedCell.Should().BeTrue();
            table[0, 1].IsMergedCell.Should().BeFalse();
            table[0, 0].TextFrame.Text.Should().Be(expectedText);
            table[1, 0].TextFrame.Text.Should().Be(expectedText);
        }
    }

    [Test]
    public void MergeCells_merges_cells()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[2].Shapes.GetById<ITable>(5);

        // Act
        table.MergeCells(table[0, 0], table[1, 0]);

        // Assert
        table[1, 0].IsMergedCell.Should().BeTrue();
    }

    [Test]
    public void MergeCells_merges_cells_with_content()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var table = pres.Slides[0].Shapes.AddTable(10, 10, 3, 4);
        table[1, 0].TextFrame.Text = "A";
        table[3, 0].TextFrame.Text = "B";

        // Act
        table.MergeCells(table[1, 0], table[2, 0]);

        // Assert
        table[1, 0].TextFrame.Text.Should().Be("A");
    }

    [Test]
    public void MergeCells_Merges0x1And1x1CellsOf3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 1], table[1, 1]);

        // Assert
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].IsMergedCell.Should().BeFalse();

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_Merges0x0To1x1RangeOf3x3Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 10);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[1, 1]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 0].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeFalse();

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 10);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 0].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_MergesMergedCellWithNonMergedCell()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 5);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[1, 1], table[1, 2]);

        // Assert
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[1, 2].IsMergedCell.Should().BeTrue();
        table[1, 1].Should().Be(table[1, 2]);
        table[3, 2].IsMergedCell.Should().BeFalse();

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 5);
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[1, 2].IsMergedCell.Should().BeTrue();
        table[1, 1].Should().Be(table[1, 2]);
        table[3, 2].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_MergesTwoMergedCells()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        ITable table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 2);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 0].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].Should().Be(table[1, 1]);
        table[0, 2].IsMergedCell.Should().BeFalse();

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 2);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 0].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].Should().Be(table[1, 1]);
        table[0, 2].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_converts_2X1_table_into_1X1_when_all_cells_are_merged()
    {
        // Arrange
        var pres = SCPresentation.Open(GetInputStream("001.pptx"));
        var table = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 3);
        var mStream = new MemoryStream();
        var totalColWidth = table.Columns[0].Width + table.Columns[1].Width;

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        table.Columns.Should().HaveCount(1);
        table.Columns[0].Width.Should().Be(totalColWidth);
        table.Rows.Should().HaveCount(1);
        table.Rows[0].Cells.Should().HaveCount(1);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        table = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 3);
        table.Columns.Should().HaveCount(1);
        table.Columns[0].Width.Should().Be(totalColWidth);
        table.Rows.Should().HaveCount(1);
        table.Rows[0].Cells.Should().HaveCount(1);
    }

    [Test]
    public void MergeCells_converts_2X2_table_into_1X1_when_all_cells_are_merged()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[2].Shapes.GetByName<ITable>("Table 5");
        var mStream = new MemoryStream();
        var mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;
        var mergedRowHeight = table.Rows[0].Height + table.Rows[1].Height;

        // Act
        table.MergeCells(table[0, 0], table[1, 1]);

        // Assert
        AssertTable(table, mergedColumnWidth, mergedRowHeight);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        table = pres.Slides[2].Shapes.GetByName<ITable>("Table 5");
        AssertTable(table, mergedColumnWidth, mergedRowHeight);

        static void AssertTable(ITable table, int expectedMergedColumnWidth, int expectedMergedRowHeight)
        {
            table.Columns.Should().HaveCount(1);
            table.Columns[0].Width.Should().Be(expectedMergedColumnWidth);
            table.Rows.Should().HaveCount(1);
            table.Rows[0].Cells.Should().HaveCount(1);
            table.Rows[0].Height.Should().Be(expectedMergedRowHeight);
        }
    }

    [Test]
    public void MergeCells_merges_0x0_And_0x1_cells_in_3x1_table()
    {
        // Arrange
        var presentation = SCPresentation.Open(GetInputStream("001.pptx"));
        var table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        var mStream = new MemoryStream();
        var mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        AssertTable(table, mergedColumnWidth);

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        AssertTable(table, mergedColumnWidth);

        void AssertTable(ITable tableSc, int expectedMergedColumnWidth)
        {
            tableSc.Columns.Should().HaveCount(2);
            tableSc.Columns[0].Width.Should().Be(expectedMergedColumnWidth);
            tableSc.Rows.Should().HaveCount(1);
            tableSc.Rows[0].Cells.Should().HaveCount(2);
        }
    }

    [Test]
    public void MergeCells_merges_0x1_and_0x2_cells()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[3].Shapes.GetById<ITable>(6);
        var mStream = new MemoryStream();
        var expectedNewColumnWidth = table.Columns[1].Width + table.Columns[2].Width;

        // Act
        table.MergeCells(table[0, 1], table[0, 2]);

        // Assert
        AssertTable(table, expectedNewColumnWidth);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        table = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 6);
        AssertTable(table, expectedNewColumnWidth);

        static void AssertTable(ITable table, int expectedNewColumnWidth)
        {
            table.Columns[1].Width.Should().Be(expectedNewColumnWidth);
            table.Rows.Should().HaveCount(1);
            table.Rows[0].Cells.Should().HaveCount(2);
        }
    }

    [Test]
    public void MergeCells_updates_columns_count()
    {
        // Arrange
        var pptx = GetInputStream("001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[3].Shapes.GetById<ITable>(6);

        // Act
        table.MergeCells(table[0, 1], table[0, 2]);

        // Assert
        table.Columns.Should().HaveCount(2);
    }

    [Test]
    public void Indexer_ReturnsCellByRowAndColumnIndexes()
    {
        // Arrange
        ITable tableCase1 =
            (ITable)SCPresentation.Open(GetInputStream("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);
        ITable tableCase2 =
            (ITable)SCPresentation.Open(GetInputStream("001.pptx")).Slides[3].Shapes.First(sp => sp.Id == 4);

        // Act
        ICell scCellCase1 = tableCase1[0, 0];
        ICell scCellCase2 = tableCase2[1, 1];

        // Assert
        scCellCase1.Should().NotBeNull();
        scCellCase2.Should().NotBeNull();
    }

    [Test]
    public void MergeCells()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slide = pres.Slides[0];
        var table = slide.Shapes.AddTable(0, 0, 3, 2);
        
        // Act
        table.MergeCells(table[0, 2], table[1, 2]);

        // Assert
        table[0, 1].Should().NotBeSameAs(table[1, 1]);
    }
    
    [Test]
    public void MergeCells_merges_0x0_and_0x1_then_1x1_and_1x2_cells()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slide = pres.Slides[0];
        var table = slide.Shapes.AddTable(0, 0, 4, 2);
        
        // Act
        table.MergeCells(table[0, 0], table[0, 1]);
        table.MergeCells(table[1, 1], table[1, 2]);
        
        // Assert
        table[1, 1].Should().BeSameAs(table[1, 2]);
    }

    [Test]
    public void MergeCells_merges_0x1_and_1x1()
    {
        // Arrange
        var pres = SCPresentation.Create();
        var slide = pres.Slides[0];
        var table = slide.Shapes.AddTable(0, 0, 4, 2);
        
        // Act
        table.MergeCells(table[0, 1], table[1, 1]);

        // Assert
        var aTableRow = table.Rows[0].ATableRow();
        aTableRow.Elements<A.TableCell>().ToList()[2].RowSpan.Should().BeNull();
    }
}