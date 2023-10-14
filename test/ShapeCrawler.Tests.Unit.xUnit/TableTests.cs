using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Drawing;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using ShapeCrawler.Tests.Unit.Helpers.Attributes;
using Xunit;
using Xunit.Abstractions;

namespace ShapeCrawler.Tests.Unit;

public class TableTests : SCTest
{
    [Xunit.Theory]
    [SlideShapeData("009_table.pptx", 3, 3, 3)]
    [SlideShapeData("001.pptx", 2, 5, 4)]
    public void Rows_Count_returns_number_of_rows(IShape shape, int expectedCount)
    {
        // Arrange
        var table = (ITable)shape;

        // Act
        var rowsCount = table.Rows.Count;

        // Assert
        rowsCount.Should().Be(expectedCount);
    }

    [Xunit.Theory]
    [MemberData(nameof(TestCasesCellIsMergedCell))]
    public void Row_Cell_IsMergedCell_returns_true_When_cell_is_merged_Vertically(ITableCell cell1, ITableCell cell2)
    {
        // Act-Assert
        cell1.IsMergedCell.Should().BeTrue();
        cell2.IsMergedCell.Should().BeTrue();
        var internalCell1 = (TableCell) cell1;
        var internalCell2 = (TableCell) cell2;
        internalCell1.RowIndex.Should().Be(internalCell2.RowIndex);
        internalCell1.ColumnIndex.Should().Be(internalCell2.ColumnIndex);
    }

    public static IEnumerable<object[]> TestCasesCellIsMergedCell()
    {
        var pptx = StreamOf("001.pptx");
        var table1 = new Presentation(pptx).Slides[1].Shapes.GetById<ITable>(3);
        yield return new object[] { table1[0, 0], table1[1, 0] };

        var pptx2 = StreamOf("001.pptx");
        var pres2 = new Presentation(pptx2);
        var table2 = pres2.Slides[1].Shapes.GetByName<ITable>("Table 5");
        yield return new object[] { table2[1, 1], table2[2, 1] };

        var pptx3 = StreamOf("001.pptx");
        var pres3 = new Presentation(pptx3);
        var table3 = pres3.Slides[3].Shapes.GetById<ITable>(4);
        yield return new object[] { table3[0, 1], table3[1, 1] };
    }

    [Xunit.Theory]
    [InlineData(0, 0, 0, 1)]
    [InlineData(0, 1, 0, 0)]
    public void MergeCells_MergesSpecifiedCellsRange(int rowIdx1, int colIdx1, int rowIdx2, int colIdx2)
    {
        // Arrange
        IPresentation presentation = new Presentation(StreamOf("001.pptx"));
        ITable table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 4);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[rowIdx1, colIdx1], table[rowIdx2, colIdx2]);

        // Assert
        table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
        table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();

        presentation.SaveAs(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 4);
        table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
        table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();
    }
}