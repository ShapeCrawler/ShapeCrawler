using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Collections;
using ShapeCrawler.Tables;
using ShapeCrawler.Tests.Helpers;
using ShapeCrawler.Tests.Properties;
using Xunit;

namespace ShapeCrawler.Tests
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    [SuppressMessage("ReSharper", "SuggestVarOrType_BuiltInTypes")]
    public class TableTests : IClassFixture<PresentationFixture>
    {
        private readonly PresentationFixture _fixture;

        public TableTests(PresentationFixture fixture)
        {
            _fixture = fixture;
        }

        [Fact]
        public void RowsAndRowCellsCounters_ReturnNumberOfRowsInTheTableAndNumberOfCellsInTheTableRow()
        {
            // Arrange
            ITable table = (ITable) _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3);

            // Act
            RowCollection tableRows = table.Rows;
            IEnumerable<ITableCell> rowCells = tableRows.First().Cells;

            // Assert
            tableRows.Should().HaveCount(3);
            rowCells.Should().HaveCount(3);
        }

        [Fact]
        public void RowRemoveAt_RemovesTableRowWithSpecifiedIndex()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._009);
            ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
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
        public void RowHeightGetter_ReturnsHeightOfTableRow()
        {
            // Arrange
            ITable table = (ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 3);

            // Act-Assert
            table.Rows[0].Height.Should().Be(370840);
        }

        [Fact]
        public void CellIsMergedCell_ReturnsTrue_WhenCellMergedWithOtherHorizontally()
        {
            // Arrange
            SCTableRow scTableRow = ((ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4)).Rows[1];
            ITableCell cell1x0 = scTableRow.Cells[0];
            ITableCell cell1x1 = scTableRow.Cells[1];

            // Act-Assert
            cell1x0.IsMergedCell.Should().BeTrue();
            cell1x1.IsMergedCell.Should().BeTrue();
            cell1x0.Should().BeSameAs(cell1x1);
        }

        [Theory]
        [MemberData(nameof(TestCasesCellIsMergedCell))]
        public void CellIsMergedCell_ReturnsTrue_WhenCellMergedWithOtherVertically(ITableCell cell1, ITableCell cell2)
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
        public void ColumnWidthGetter_ReturnsTableColumnWidthInEMU()
        {
            // Arrange
            ITable table = (ITable)_fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4);
            Column column = table.Columns[0];

            // Act
            long columnWidth = column.Width;

            // Assert
            columnWidth.Should().Be(3505199);
        }

        [Fact]
        public void ColumnWidthSetter_ChangeTableColumnWidth()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._001);
            ITable table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 3);
            const long newColumnWidth = 4074000;
            var mStream = new MemoryStream();

            // Act
            table.Columns[0].Width = newColumnWidth;

            // Assert
            table.Columns[0].Width.Should().Be(newColumnWidth);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 3);
            table.Columns[0].Width.Should().Be(newColumnWidth);
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
            table[0, 0].TextBox.Text.Should().Be($"id5{Environment.NewLine}Text0_1");

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5);
            table[0, 0].IsMergedCell.Should().BeTrue();
            table[0, 1].IsMergedCell.Should().BeTrue();
            table[0, 0].TextBox.Text.Should().Be($"id5{Environment.NewLine}Text0_1");
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
                tableSc[0, 1].TextBox.Text.Should().Be("Text0_2");
                tableSc[0, 2].TextBox.Text.Should().Be("Text0_2");
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
                table[0, 0].TextBox.Text.Should().Be(expectedText);
                table[1, 0].TextBox.Text.Should().Be(expectedText);
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
        public void MergeCells_Converts2X1TableInto1X1_WhenAllCellsAreMerged()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._001);
            ITable table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 3);
            var mStream = new MemoryStream();
            long totalColWidth = table.Columns[0].Width + table.Columns[1].Width;

            // Act
            table.MergeCells(table[0, 0], table[0, 1]);

            // Assert
            table.Columns.Should().HaveCount(1);
            table.Columns[0].Width.Should().Be(totalColWidth);
            table.Rows.Should().HaveCount(1);
            table.Rows[0].Cells.Should().HaveCount(1);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 3);
            table.Columns.Should().HaveCount(1);
            table.Columns[0].Width.Should().Be(totalColWidth);
            table.Rows.Should().HaveCount(1);
            table.Rows[0].Cells.Should().HaveCount(1);
        }

        [Fact(DisplayName = "MergeCells #9")]
        public void MergeCells_Converts2X2TableInto1X1_WhenAllCellsAreMerged()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._001);
            ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5) ;
            var mStream = new MemoryStream();
            long mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;
            long mergedRowHeight = table.Rows[0].Height + table.Rows[1].Height;

            // Act
            table.MergeCells(table[0, 0], table[1,1]);

            // Assert
            AssertTable(table, mergedColumnWidth, mergedRowHeight);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5) ;
            AssertTable(table, mergedColumnWidth, mergedRowHeight);

            static void AssertTable(ITable tableSc, long expectedMergedColumnWidth, long expectedMergedRowHeight)
            {
                tableSc.Columns.Should().HaveCount(1);
                tableSc.Columns[0].Width.Should().Be(expectedMergedColumnWidth);
                tableSc.Rows.Should().HaveCount(1);
                tableSc.Rows[0].Cells.Should().HaveCount(1);
                tableSc.Rows[0].Height.Should().Be(expectedMergedRowHeight);
            }
        }

        [Fact(DisplayName = "MergeCells #10")]
        public void MergeCells_Merges0x0and0x1CellsIn3x1Table()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._001);
            ITable table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6) ;
            var mStream = new MemoryStream();
            long mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;

            // Act
            table.MergeCells(table[0, 0], table[0, 1]);

            // Assert
            AssertTable(table, mergedColumnWidth);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6) ;
            AssertTable(table, mergedColumnWidth);

            void AssertTable(ITable tableSc, long expectedMergedColumnWidth)
            {
                tableSc.Columns.Should().HaveCount(2);
                tableSc.Columns[0].Width.Should().Be(expectedMergedColumnWidth);
                tableSc.Rows.Should().HaveCount(1);
                tableSc.Rows[0].Cells.Should().HaveCount(2);
            }
        }

        [Fact(DisplayName = "MergeCells #11")]
        public void MergeCells_Merges0x1and0x2CellsIn3x1Table()
        {
            // Arrange
            IPresentation presentation = SCPresentation.Open(Resources._001);
            ITable table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6) ;
            var mStream = new MemoryStream();
            long mergedColumnWidth = table.Columns[1].Width + table.Columns[2].Width;

            // Act
            table.MergeCells(table[0, 1], table[0, 2]);

            // Assert
            AssertTable(table, mergedColumnWidth);

            presentation.SaveAs(mStream);
            presentation = SCPresentation.Open(mStream);
            table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6) ;
            AssertTable(table, mergedColumnWidth);

            static void AssertTable(ITable tableSc, long expectedMergedColumnWidth)
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
            ITableCell scCellCase1 = tableCase1[0, 0];
            ITableCell scCellCase2 = tableCase2[1, 1];

            // Assert
            scCellCase1.Should().NotBeNull();
            scCellCase2.Should().NotBeNull();
        }
    }
}
