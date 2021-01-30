using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Tables;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Properties;
using Xunit;

namespace ShapeCrawler.Tests.Unit
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
            TableSc table = _fixture.Pre009.Slides[2].Shapes.First(sp => sp.Id == 3).Table;

            // Act
            RowCollection tableRows = table.Rows;
            IEnumerable<CellSc> rowCells = tableRows.First().Cells;

            // Assert
            tableRows.Should().HaveCount(3);
            rowCells.Should().HaveCount(3);
        }

        [Fact]
        public void RowRemoveAt_RemovesTableRowWithSpecifiedIndex()
        {
            // Arrange
            PresentationSc presentation = PresentationSc.Open(Resources._009, true);
            TableSc table = presentation.Slides[2].Shapes.First(sp => sp.Id == 3).Table;
            int originRowsCount = table.Rows.Count;
            var mStream = new MemoryStream();

            // Act
            table.Rows.RemoveAt(0);

            // Assert
            table.Rows.Should().HaveCountLessThan(originRowsCount);

            presentation.SaveAs(mStream);
            presentation.Close();
            table = PresentationSc.Open(mStream, false).Slides[2].Shapes.First(sp => sp.Id == 3).Table;
            table.Rows.Should().HaveCountLessThan(originRowsCount);
        }

        [Fact]
        public void CellIsMergedCell_ReturnsTrueWhenTheCellBelongToMergedCellsGroup()
        {
            // Arrange
            CellSc tableCellSc = _fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4).Table.Rows[1].Cells[0];

            // Act
            bool isMergedCell = tableCellSc.IsMergedCell;

            // Assert
            isMergedCell.Should().BeTrue();
        }

        // TODO: Add test case - merging two already merged cells; merging merged cell with un-merged cell
        [Theory(Skip = "In Progress")]
        [InlineData(0, 0, 0, 1)]
        [InlineData(0, 1, 0, 0)]
        public void MergeCells_MergesSpecifiedCellsRange(int rowIdx1, int colIdx1, int rowIdx2, int colIdx2)
        {
            // Arrange
            PresentationSc presentation = PresentationSc.Open(Resources._001, true);
            TableSc table = presentation.Slides[1].Shapes.First(sp => sp.Id == 4).Table;
            var mStream = new MemoryStream();

            // Act
            //table.MergeCells(table[rowIdx1, colIdx1], table[rowIdx2, colIdx2]);

            // Assert
            table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
            table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();

            presentation.SaveAs(mStream);
            presentation.Close();
            presentation = PresentationSc.Open(mStream, false);
            table = presentation.Slides[1].Shapes.First(sp => sp.Id == 4).Table;
            table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
            table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();
        }

        [Fact]
        public void Indexer_ReturnsCellByRowAndColumnIndexes()
        {
            // Arrange
            TableSc table = _fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4).Table;

            // Act
            CellSc tableCell = table[0, 0];

            // Assert
            tableCell.Should().NotBeNull();
        }
    }
}
