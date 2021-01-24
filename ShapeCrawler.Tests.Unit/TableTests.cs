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
            IEnumerable<Cell> rowCells = tableRows.First().Cells;

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
            Cell tableCell = _fixture.Pre001.Slides[1].Shapes.First(sp => sp.Id == 4).Table.Rows[1].Cells[0];

            // Act
            bool isMergedCell = tableCell.IsMergedCell;

            // Assert
            isMergedCell.Should().BeTrue();
        }
    }
}
