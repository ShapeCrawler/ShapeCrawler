using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using ShapeCrawler.Enums;
using ShapeCrawler.Models;
using ShapeCrawler.Models.SlideComponents;
using ShapeCrawler.Tables;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit;

namespace ShapeCrawler.Tests.Unit
{
    [SuppressMessage("ReSharper", "SuggestVarOrType_SimpleTypes")]
    public class TableTests : IClassFixture<PptxFixture>
    {
        private readonly PptxFixture _fixture;

        public TableTests(PptxFixture fixture)
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
    }
}
