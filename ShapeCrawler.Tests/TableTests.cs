using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
using Xunit;

namespace ShapeCrawler.Tests;

[SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
[SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
public class TableTests : ShapeCrawlerTest, IClassFixture<PresentationFixture>
{
    private readonly PresentationFixture _fixture;

    public TableTests(PresentationFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Rows_Count_returns_number_of_rows()
    {
        // Arrange
        var table = (ITable)this._fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3);

        // Act
        var rowsCount = table.Rows.Count;

        // Assert
        rowsCount.Should().Be(3);
    }
    
    [Fact]
    public void Rows_RemoveAt_removes_row_with_specified_index()
    {
        // Arrange
        var presentation = SCPresentation.Open(Resources._009);
        var table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        int originRowsCount = table.Rows.Count;
        var mStream = new MemoryStream();

        // Act
        table.Rows.RemoveAt(0);

        // Assert
        table.Rows.Should().HaveCountLessThan(originRowsCount);
        presentation.SaveAs(mStream);
        table = (ITable)SCPresentation.Open(mStream).Slides[2].Shapes.First(sp => sp.Id == 3);
        table.Rows.Should().HaveCountLessThan(originRowsCount);
    }

    [Fact]
    public void Rows_Add_adds_row()
    {
        // Arrange
        var pptx = GetTestStream("table-case001.pptx");
        var pres = SCPresentation.Open(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");

        // Act
        table.Rows.Add();
        
        // Assert
        table.Rows.Should().HaveCount(2);
        var errors = PptxValidator.Validate(pres);
        errors.Should().BeEmpty();
    }
    
    [Fact]
    public void Row_Cells_Count_returns_number_of_cells_in_the_row()
    {
        // Arrange
        var table = (ITable)this._fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3);

        // Act
        var cellsCount = table.Rows[0].Cells.Count;

        // Assert
        cellsCount.Should().Be(3);
    }

    [Fact]
    public void Row_Height_Getter_returns_height_of_row()
    {
        // Arrange
        var table = (ITable)this._fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 3);
        
        // Act
        var rowHeight = table.Rows[0].Height; 

        // Act-Assert
        rowHeight.Should().Be(370840);
    }

    [Theory]
    [MemberData(nameof(TestCasesCellIsMergedCell))]
    public void CellIsMergedCell_ReturnsTrue_WhenCellMergedWithOtherVertically(ICell cell1, ICell cell2)
    {
        cell1.IsMergedCell.Should().BeTrue();
        cell2.IsMergedCell.Should().BeTrue();
        cell1.Should().Be(cell2);
    }

    public static IEnumerable<object[]> TestCasesCellIsMergedCell()
    {
        ITable table = (ITable)SCPresentation.Open(Resources._001).Slides[1].Shapes.First(sp => sp.Id == 3);
        yield return new object[] {table[0, 0], table[1, 0]};

        table = (ITable)SCPresentation.Open(Resources._001).Slides[1].Shapes.First(sp => sp.Id == 5);
        yield return new object[] { table[1, 1], table[2, 1] };

        table = (ITable)SCPresentation.Open(Resources._001).Slides[3].Shapes.First(sp => sp.Id == 4);
        yield return new object[] { table[0, 1], table[1, 1] };
    }

    [Fact]
    public void ColumnsCount_ReturnsNumberOfColumnsInTheTable()
    {
        // Arrange
        ITable table = (ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4);

        // Act
        int columnsCount = table.Columns.Count;

        // Assert
        columnsCount.Should().Be(3);
    }

    [Fact]
    public void Column_Width_Getter_returns_width_of_column_in_pixels()
    {
        // Arrange
        var table = (ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4);

        // Act
        var columnWidth = table.Columns[0].Width;

        // Assert
        columnWidth.Should().Be(367);
    }

    [Fact]
    public void Column_Width_Setter_sets_width_of_column()
    {
        // Arrange
        var pres = SCPresentation.Open(Resources._001);
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
    
    [Fact]
    public void Row_Cell_IsMergedCell_returns_True_When_cell_is_merged_with_other_horizontally()
    {
        // Arrange
        var row = ((ITable)this._fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4)).Rows[1];
        var cell1x0 = row.Cells[0];
        var cell1x1 = row.Cells[1];

        // Act
        var isMerged1 = cell1x0.IsMergedCell;
        var isMerged2 = cell1x1.IsMergedCell; 
        
        // Act-Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        cell1x0.Should().BeSameAs(cell1x1);
    }
    
    [Fact]
    public void Row_Clone_cloning_row_increases_row_count_by_one()
    {
        // Arrange
        var pres = SCPresentation.Open(Resources.tables_case001);
        var targetTable = pres.Slides.First().Shapes.OfType<ITable>().FirstOrDefault();
        var rowCountBefore = targetTable.Rows.Count;
        var row = targetTable.Rows.Last(); 

        // Act
        row.Clone();

        // Assert
        var rowCountAfter = targetTable.Rows.Count;
        rowCountAfter.Should().Be(rowCountBefore + 1);
    }

#if DEBUG

    [Theory]
    [InlineData(0, 0, 0, 1)]
    [InlineData(0, 1, 0, 0)]
    public void MergeCells_MergesSpecifiedCellsRange(int rowIdx1, int colIdx1, int rowIdx2, int colIdx2)
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
        ITable table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 4);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[rowIdx1, colIdx1], table[rowIdx2, colIdx2]);

        // Assert
        table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
        table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 4);
        table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
        table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();
    }

    [Fact(DisplayName = "MergeCells #1")]
    public void MergeCells_Merges0x0And0x1CellsOf2x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact(DisplayName = "MergeCells #2")]
    public void MergeCells_Merges0x1And0x2CellsOf3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact(DisplayName = "MergeCells #3")]
    public void MergeCells_Merges0x0And0x1And0x2CellsOf3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact(DisplayName = "MergeCells #4")]
    public void MergeCells_Merges0x0And0x1MergedCellsWith0x2CellIn3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact(DisplayName = "MergeCells #5")]
    public void MergeCells_Merges0x0And1x0CellsOf2x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[1, 0]);

        // Assert
        AssertTable(table);

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5);
        AssertTable(table);

        void AssertTable(ITable table)
        {
            string expectedText = $"id5{Environment.NewLine}Text1_0";
            table[0, 0].IsMergedCell.Should().BeTrue();
            table[1, 0].IsMergedCell.Should().BeTrue();
            table[0, 1].IsMergedCell.Should().BeFalse();
            table[0, 0].TextFrame.Text.Should().Be(expectedText);
            table[1, 0].TextFrame.Text.Should().Be(expectedText);
        }
    }

    [Fact(DisplayName = "MergeCells #6")]
    public void MergeCells_Merges0x1And1x1CellsOf3x2Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact(DisplayName = "MergeCells #7")]
    public void MergeCells_Merges0x0To1x1RangeOf3x3Table()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact]
    public void MergeCells_MergesMergedCellWithNonMergedCell()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact]
    public void MergeCells_MergesTwoMergedCells()
    {
        // Arrange
        IPresentation presentation = SCPresentation.Open(Resources._001);
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

    [Fact(DisplayName = "MergeCells #8")]
    public void MergeCells_converts_2X1_table_into_1X1_when_all_cells_are_merged()
    {
        // Arrange
        var pres = SCPresentation.Open(Resources._001);
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

    [Fact(DisplayName = "MergeCells #9")]
    public void MergeCells_converts_2X2_table_into_1X1_when_all_cells_are_merged()
    {
        // Arrange
        var pres = SCPresentation.Open(Resources._001);
        var table = (ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 5) ;
        var mStream = new MemoryStream();
        var mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;
        var mergedRowHeight = table.Rows[0].Height + table.Rows[1].Height;

        // Act
        table.MergeCells(table[0, 0], table[1,1]);

        // Assert
        AssertTable(table, mergedColumnWidth, mergedRowHeight);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        table = (ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 5) ;
        AssertTable(table, mergedColumnWidth, mergedRowHeight);

        static void AssertTable(ITable tableSc, int expectedMergedColumnWidth, long expectedMergedRowHeight)
        {
            tableSc.Columns.Should().HaveCount(1);
            tableSc.Columns[0].Width.Should().Be(expectedMergedColumnWidth);
            tableSc.Rows.Should().HaveCount(1);
            tableSc.Rows[0].Cells.Should().HaveCount(1);
            tableSc.Rows[0].Height.Should().Be(expectedMergedRowHeight);
        }
    }

    [Fact(DisplayName = "MergeCells #10")]
    public void MergeCells_merges_0x0_And_0x1_cells_in_3x1_table()
    {
        // Arrange
        var presentation = SCPresentation.Open(Resources._001);
        var table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6) ;
        var mStream = new MemoryStream();
        var mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        AssertTable(table, mergedColumnWidth);

        presentation.SaveAs(mStream);
        presentation = SCPresentation.Open(mStream);
        table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6) ;
        AssertTable(table, mergedColumnWidth);

        void AssertTable(ITable tableSc, int expectedMergedColumnWidth)
        {
            tableSc.Columns.Should().HaveCount(2);
            tableSc.Columns[0].Width.Should().Be(expectedMergedColumnWidth);
            tableSc.Rows.Should().HaveCount(1);
            tableSc.Rows[0].Cells.Should().HaveCount(2);
        }
    }

    [Fact(DisplayName = "MergeCells #11")]
    public void MergeCells_merges_0x1_And_0x2_cells_in_3x1_table()
    {
        // Arrange
        var pres = SCPresentation.Open(Resources._001);
        var table = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 6) ;
        var mStream = new MemoryStream();
        var mergedColumnWidth = table.Columns[1].Width + table.Columns[2].Width;

        // Act
        table.MergeCells(table[0, 1], table[0, 2]);

        // Assert
        AssertTable(table, mergedColumnWidth);

        pres.SaveAs(mStream);
        pres = SCPresentation.Open(mStream);
        table = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 6) ;
        AssertTable(table, mergedColumnWidth);

        static void AssertTable(ITable tableSc, int expectedMergedColumnWidth)
        {
            tableSc.Columns.Should().HaveCount(2);
            tableSc.Columns[1].Width.Should().Be(expectedMergedColumnWidth);
            tableSc.Rows.Should().HaveCount(1);
            tableSc.Rows[0].Cells.Should().HaveCount(2);
        }
    }
#endif

    [Fact]
    public void Indexer_ReturnsCellByRowAndColumnIndexes()
    {
        // Arrange
        ITable tableCase1 = (ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4);
        ITable tableCase2 = (ITable)_fixture.Pre001.Slides[3].Shapes.First(sp => sp.Id == 4);

        // Act
        ICell scCellCase1 = tableCase1[0, 0];
        ICell scCellCase2 = tableCase2[1, 1];

        // Assert
        scCellCase1.Should().NotBeNull();
        scCellCase2.Should().NotBeNull();
    }
}