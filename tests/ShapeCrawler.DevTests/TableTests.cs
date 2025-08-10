using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.DevTests.Helpers;
using ShapeCrawler.Tables;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.DevTests;

public class TableTests : SCTest
{
    [Test]
    public void TableStyle_Getter_return_style_of_table()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var table = pres.Slide(3).Shape("Таблица 4").Table;

        // Act 
        var tableStyle = table.TableStyle;

        // Assert
        tableStyle.Should().BeEquivalentTo(CommonTableStyles.MediumStyle2Accent1);
        pres.Validate();
    }

    [Test]
    public void TableStyle_Setter_sets_style()
    {
        // Arrange
        var pres = new Presentation(TestAsset("009_table.pptx"));
        var table = pres.Slide(3).Shape("Таблица 4").Table;
        var mStream = new MemoryStream();

        // Act
        table.TableStyle = CommonTableStyles.ThemedStyle1Accent4;

        // Assert
        table.TableStyle.Should().BeEquivalentTo(CommonTableStyles.ThemedStyle1Accent4);
        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = pres.Slide(3).Shape("Таблица 4").Table;
        table.TableStyle.Should().BeEquivalentTo(CommonTableStyles.ThemedStyle1Accent4);
        pres.Validate();
    }

    [Test]
    public void CellBorder_Getter_return_bottom_border_color()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;

        // Act 
        var tableCell = table[1, 2];

        // Assert
        tableCell.BottomBorder.Color.Should().Be("FF0000");
    }

    [Test]
    public void CellBorder_Setter_right_border_color()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;
        var mStream = new MemoryStream();

        // Act
        table[1, 2].RightBorder.Color = "00FF00";

        // Assert
        table[1, 2].RightBorder.Color.Should().Be("00FF00");
        pres.Save(mStream);

        pres = new Presentation(mStream);
        table = pres.Slide(1).Shape("Table 1").Table;
        table[1, 2].RightBorder.Color.Should().Be("00FF00");
        pres.Validate();
    }

    [Test]
    public void CellMargins_Getter_return_cell_margin()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Shape("Table 2").Table;

        // Assert
        var tableCell = table[0, 0];
        tableCell.TextBox.TopMargin.Should().BeApproximately(28.34m, 0.01m);
        tableCell.TextBox.RightMargin.Should().BeApproximately(28.34m, 0.01m);
        tableCell.TextBox.LeftMargin.Should().BeApproximately(28.34m, 0.01m);
        tableCell.TextBox.BottomMargin.Should().BeApproximately(28.34m, 0.01m);
    }

    [Test]
    public void CellMargins_Setter_cell_margin()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Shape("Table 2").Table;
        var mStream = new MemoryStream();

        // Act
        var tableCell = table[0, 0];
        tableCell.TextBox.LeftMargin = 10m;
        tableCell.TextBox.RightMargin = 10m;
        tableCell.TextBox.TopMargin = 20m;
        tableCell.TextBox.BottomMargin = 20m;

        // Assert
        tableCell.TextBox.TopMargin.Should().Be(20m);
        pres.Save(mStream);

        pres = new Presentation(mStream);
        table = pres.Slide(1).Shape("Table 2").Table;

        tableCell = table[0, 0];
        tableCell.TextBox.LeftMargin.Should().Be(10m);
        tableCell.TextBox.RightMargin.Should().Be(10m);
        tableCell.TextBox.TopMargin.Should().Be(20m);
        tableCell.TextBox.BottomMargin.Should().Be(20m);

        pres.Validate();
    }

    [Test]
    public void CellTextAlignment_Getter_return_cell_alignment()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Shape("Table 2").Table;

        // Assert
        var tableCell = table[3, 2];
        tableCell.TextBox.VerticalAlignment.Should().Be(TextVerticalAlignment.Bottom);
        tableCell.TextBox.Paragraphs[0].HorizontalAlignment.Should().Be(TextHorizontalAlignment.Right);
    }

    [Test]
    public void CellTextAlignment_Setter_cell_alignment()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Shape("Table 2").Table;
        var mStream = new MemoryStream();

        // Act
        var tableCell = table[1, 1];
        tableCell.TextBox.VerticalAlignment = TextVerticalAlignment.Middle;
        tableCell.TextBox.Paragraphs[0].HorizontalAlignment = TextHorizontalAlignment.Center;

        // Assert
        tableCell.TextBox.VerticalAlignment.Should().Be(TextVerticalAlignment.Middle);
        tableCell.TextBox.Paragraphs[0].HorizontalAlignment.Should().Be(TextHorizontalAlignment.Center);
        pres.Save(mStream);

        pres = new Presentation(mStream);
        table = pres.Slide(1).Shape("Table 2").Table;

        tableCell = table[1, 1];
        tableCell.TextBox.VerticalAlignment.Should().Be(TextVerticalAlignment.Middle);
        tableCell.TextBox.Paragraphs[0].HorizontalAlignment.Should().Be(TextHorizontalAlignment.Center);

        pres.Validate();
    }
    
    [Test]
    public void Rows_RemoveAt_removes_row_with_specified_index()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var pres = new Presentation(pptx);
        var table = (ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 3);
        int originRowsCount = table.Rows.Count;
        var mStream = new MemoryStream();

        // Act
        table.Rows.RemoveAt(0);

        // Assert
        table.Rows.Should().HaveCountLessThan(originRowsCount);
        pres.Save(mStream);
        table = (ITable)new Presentation(mStream).Slides[2].Shapes.First(sp => sp.Id == 3);
        table.Rows.Should().HaveCountLessThan(originRowsCount);
    }

    [Test]
    public void Rows_Add_adds_row()
    {
        // Arrange
        var pptx = TestAsset("table-case001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(1).Shape("Table 1").Table;

        // Act
        table.Rows.Add();

        // Assert
        table.Rows.Should().HaveCount(2);
        pres.Validate();
    }

    [Test]
    public void Row_Cells_Count_returns_number_of_cells_in_the_row()
    {
        // Arrange
        var pptx = TestAsset("009_table.pptx");
        var table = (ITable)new Presentation(pptx).Slides[2].Shapes
            .First(sp => sp.Id == 3);

        // Act
        var cellsCount = table.Rows[0].Cells.Count;

        // Assert
        cellsCount.Should().Be(3);
    }

    [Test]
    public void Row_Height_Getter()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = (ITable)pres.Slide(2).Shapes.First(sp => sp.Id == 3);

        // Act & Assert
        table.Rows[0].Height.Should().Be(29.2m);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_True_When_cell_is_merged()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var row = pres.Slide(2).Shape("Table 4").Table.Rows[1];
        var cell1X0 = row.Cells[0];
        var cell1X1 = row.Cells[1];

        // Act
        var isMerged1 = cell1X0.IsMergedCell;
        var isMerged2 = cell1X1.IsMergedCell;

        // Act-Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();

        var cell1 = (TableCell)cell1X0;
        var cell2 = (TableCell)cell1X1;
        cell1.RowIndex.Should().Be(cell2.RowIndex);
        cell1.ColumnIndex.Should().Be(cell2.ColumnIndex);
    }

    [Test]
    public void Row_Cell_TopBorder_Width_Setter_sets_top_border_width_in_points()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(40, 40, 2, 1);
        var table = (ITable)slide.Shapes.Last();
        var cell = table[0, 1];

        // Act
        cell.TopBorder.Width = 2;

        // Assert
        cell.TopBorder.Width.Should().Be(2);
    }

    [Test]
    public void Row_Cell_LeftBorder_Width_Setter_sets_left_border_width_in_points()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var cell = pres.Slide(1).Shape("Table 1").Table[0, 0];

        // Act
        cell.LeftBorder.Width = 2;

        // Assert
        cell.LeftBorder.Width.Should().Be(2);
    }

    [Test]
    public void Row_Cell_TopBorder_Width_Getter_returns_top_border_width_in_points()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(40, 40, 2, 1);
        var table = (ITable)slide.Shapes.Last();
        var cell = table[0, 1];

        // Act-Assert
        cell.TopBorder.Width.Should().Be(1);
    }

    [Test]
    public void Row_Clone_cloning_row_increases_row_count_by_one()
    {
        // Arrange
        var pptx = TestAsset("065 table.pptx");
        var pres = new Presentation(pptx);
        var targetTable = pres.Slides.First().Shapes.OfType<ITable>().FirstOrDefault();
        var rowCountBefore = targetTable.Rows.Count;
        var row = targetTable.Rows.Last();

        // Act
        row.Duplicate();

        // Assert
        var rowCountAfter = targetTable.Rows.Count;
        rowCountAfter.Should().Be(rowCountBefore + 1);
    }

    [Test]
    public void Row_Height_Setter()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var tableShape = pres.Slide(1).Shape("Table 1");
        var row = tableShape.Table.Rows[0];

        // Act
        row.Height = 39;

        // Assert
        row.Height.Should().Be(39);
        tableShape.Height.Should().Be(39);
    }

    [Test]
    public void MergeCells_merges_0x0_and_0x1_cells_of_2x2_table()
    {
        // Arrange
        IPresentation presentation = new Presentation(TestAsset("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].TextBox.Text.Should().Be($"id5{Environment.NewLine}Text0_1");

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 5);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].TextBox.Text.Should().Be($"id5{Environment.NewLine}Text0_1");
    }

    [Test]
    public void MergeCells_merges_0x1_and_0x2_cells_of_3x2_table()
    {
        // Arrange
        IPresentation presentation = new Presentation(TestAsset("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 1], table[0, 2]);

        // Assert
        AssertTable(table);
        presentation.Save(mStream);
        presentation = new Presentation(mStream);
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

    [Test]
    public void MergeCells_merges0x0_and_0x1_and_0x2_cells_of_3x2_table()
    {
        // Arrange
        IPresentation presentation = new Presentation(TestAsset("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 2]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();
    }

    [Test]
    public void MergeCells_merges_0x0_and_0x1_merged_cells_with_0x2_cell_of_3x2_table()
    {
        // Arrange
        IPresentation presentation = new Presentation(TestAsset("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 2]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 7);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeTrue();
    }

    [Test]
    public void MergeCells_merges_0x0_and_1x0_cells_of_2x2_table()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(2).Shape(5).Table;
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[1, 0]);

        // Assert
        AssertTable(table);
        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = (ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 5);
        AssertTable(table);

        void AssertTable(ITable table)
        {
            string expectedText = $"id5{Environment.NewLine}Text1_0";
            table[0, 0].IsMergedCell.Should().BeTrue();
            table[0, 1].IsMergedCell.Should().BeFalse();
            table[0, 0].TextBox.Text.Should().Be(expectedText);
            table[1, 0].TextBox.Text.Should().Be(expectedText);
        }
    }

    [Test]
    public void MergeCells_merges_cells()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(2).Shape(5).Table;

        // Act
        table.MergeCells(table[0, 0], table[1, 0]);

        // Assert
        table[1, 0].IsMergedCell.Should().BeTrue();
    }

    [Test]
    public void MergeCells_merges_cells_with_content()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        pres.Slides[0].Shapes.AddTable(10, 10, 3, 4);
        var table = (ITable)pres.Slides[0].Shapes.Last();
        table[1, 0].TextBox.SetText("A");
        table[3, 0].TextBox.SetText("B");

        // Act
        table.MergeCells(table[1, 0], table[2, 0]);

        // Assert
        table[1, 0].TextBox.Text.Should().Be("A");
    }

    [Test]
    public void MergeCells_merges_0x1_and_1x1_cells_of_3x2_table()
    {
        // Arrange
        IPresentation presentation = new Presentation(TestAsset("001.pptx"));
        ITable table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 1], table[1, 1]);

        // Assert
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].IsMergedCell.Should().BeFalse();

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 3);
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 0].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_merges_0x0_to_1x1_range_of_3x3_table()
    {
        // Arrange
        IPresentation presentation = new Presentation(TestAsset("001.pptx"));
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

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[2].Shapes.First(sp => sp.Id == 10);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 0].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_merges_merged_cell_with_non_merged_cell()
    {
        // Arrange
        IPresentation presentation = new Presentation(TestAsset("001.pptx"));
        ITable table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 5);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[1, 1], table[1, 2]);

        // Assert
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[1, 2].IsMergedCell.Should().BeTrue();

        var cell1 = (TableCell)table[1, 1];
        var cell2 = (TableCell)table[1, 2];
        cell1.RowIndex.Should().Be(cell2.RowIndex);
        cell1.ColumnIndex.Should().Be(cell2.ColumnIndex);

        table[3, 2].IsMergedCell.Should().BeFalse();

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 5);
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[1, 2].IsMergedCell.Should().BeTrue();

        cell1 = (TableCell)table[1, 1];
        cell2 = (TableCell)table[1, 2];
        cell1.RowIndex.Should().Be(cell2.RowIndex);
        cell1.ColumnIndex.Should().Be(cell2.ColumnIndex);

        table[3, 2].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_merges_two_merged_cells()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slide(4).Shape("Table 2").Table;
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 0].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();

        var cell_0_0 = (TableCell)table[0, 0];
        var cell_1_1 = (TableCell)table[1, 1];
        cell_0_0.RowIndex.Should().Be(cell_1_1.RowIndex);
        cell_0_0.ColumnIndex.Should().Be(cell_1_1.ColumnIndex);

        table[0, 2].IsMergedCell.Should().BeFalse();
        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 2);
        table[0, 0].IsMergedCell.Should().BeTrue();
        table[0, 1].IsMergedCell.Should().BeTrue();
        table[1, 0].IsMergedCell.Should().BeTrue();
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[0, 2].IsMergedCell.Should().BeFalse();
    }

    [Test]
    public void MergeCells_converts_2x1_table_into_1x1_when_all_cells_are_merged()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
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

        pres.Save(mStream);
        pres = new Presentation(mStream);
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
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = (ITable)pres.Slides[2].Shapes.First(sp => sp.Id == 5);
        var mStream = new MemoryStream();
        var mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;
        var mergedRowHeight = table.Rows[0].Height + table.Rows[1].Height;

        // Act
        table.MergeCells(table[0, 0], table[1, 1]);

        // Assert
        AssertTable(table, mergedColumnWidth, mergedRowHeight);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = pres.Slide(3).Shape("Table 5").Table;
        AssertTable(table, mergedColumnWidth, mergedRowHeight);

        static void AssertTable(ITable table, decimal expectedMergedColumnWidth, decimal expectedMergedRowHeight)
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
        var presentation = new Presentation(TestAsset("001.pptx"));
        var table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        var mStream = new MemoryStream();
        var mergedColumnWidth = table.Columns[0].Width + table.Columns[1].Width;

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);

        // Assert
        AssertTable(table, mergedColumnWidth);

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[3].Shapes.First(sp => sp.Id == 6);
        AssertTable(table, mergedColumnWidth);

        void AssertTable(ITable tableSc, decimal expectedMergedColumnWidth)
        {
            tableSc.Columns.Should().HaveCount(2);
            tableSc.Columns[0].Width.Should().BeApproximately(expectedMergedColumnWidth, 0.01m);
            tableSc.Rows.Should().HaveCount(1);
            tableSc.Rows[0].Cells.Should().HaveCount(2);
        }
    }

    [Test]
    public void MergeCells_merges_0x1_and_0x2_cells()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(4).Shape(6).Table;
        var mStream = new MemoryStream();
        var expectedNewColumnWidth = table.Columns[1].Width + table.Columns[2].Width;

        // Act
        table.MergeCells(table[0, 1], table[0, 2]);

        // Assert
        AssertTable(table, expectedNewColumnWidth);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 6);
        AssertTable(table, expectedNewColumnWidth);

        static void AssertTable(ITable table, decimal expectedNewColumnWidth)
        {
            table.Columns[1].Width.Should().BeApproximately(expectedNewColumnWidth, 0.01m);
            table.Rows.Should().HaveCount(1);
            table.Rows[0].Cells.Should().HaveCount(2);
        }
    }

    [Test]
    public void MergeCells_updates_columns_count()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slide(4).Shape(6).Table;

        // Act
        table.MergeCells(table[0, 1], table[0, 2]);

        // Assert
        table.Columns.Should().HaveCount(2);
    }

    [Test]
    public void Indexer_returns_cell_by_row_and_column_indexes()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var tableCase1 = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 4);
        var tableCase2 = (ITable)pres.Slides[3].Shapes.First(sp => sp.Id == 4);

        // Act
        var cell1 = tableCase1[0, 0];
        var cell2 = tableCase2[1, 1];

        // Assert
        cell1.Should().NotBeNull();
        cell2.Should().NotBeNull();
    }

    [Test]
    public void MergeCells()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 3, 2);
        var table = slide.Shapes.Last() as ITable;

        // Act
        table.MergeCells(table[0, 2], table[1, 2]);

        // Assert
        table[0, 1].Should().NotBeSameAs(table[1, 1]);
    }

    [Test]
    public void MergeCells_merges_0x0_and_0x1_then_1x1_and_1x2_cells()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 4, 2);
        var table = slide.Shapes.Last() as ITable;

        // Act
        table.MergeCells(table[0, 0], table[0, 1]);
        table.MergeCells(table[1, 1], table[1, 2]);

        // Assert
        var cell1 = (TableCell)table[1, 1];
        var cell2 = (TableCell)table[1, 2];
        cell1.RowIndex.Should().Be(cell2.RowIndex);
        cell1.ColumnIndex.Should().Be(cell2.ColumnIndex);
    }

    [Test]
    public void MergeCells_merges_0x1_and_1x1()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 4, 2);
        var table = (ITable)slide.Shapes.Last();

        // Act
        table.MergeCells(table[0, 1], table[1, 1]);

        // Assert
        var aTableRow = (TableRow)table.Rows[0];
        aTableRow.ATableRow.Elements<A.TableCell>().ToList()[2].RowSpan.Should().BeNull();
    }

    [Test]
    [SlideShape("009_table.pptx", 3, 3, 3)]
    [SlideShape("001.pptx", 2, 5, 4)]
    public void Rows_Count_returns_number_of_rows(IShape shape, int expectedCount)
    {
        // Arrange
        var table = (ITable)shape;

        // Act
        var rowsCount = table.Rows.Count;

        // Assert
        rowsCount.Should().Be(expectedCount);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_true_When_0x0_and_1x0_cells_are_merged()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slide(2).Shape(3).Table;
        var cell1 = table[0, 0];
        var cell2 = table[1, 0];

        // Act
        var isMerged1 = cell1.IsMergedCell;
        var isMerged2 = cell2.IsMergedCell;

        // Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        var internalCell1 = (TableCell)cell1;
        var internalCell2 = (TableCell)cell2;
        internalCell1.RowIndex.Should().Be(internalCell2.RowIndex);
        internalCell1.ColumnIndex.Should().Be(internalCell2.ColumnIndex);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_true_When_1x1_and_2x1_cells_are_merged()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slide(2).Shape("Table 5").Table;
        var cell1 = table[1, 1];
        var cell2 = table[2, 1];

        // Act
        var isMerged1 = cell1.IsMergedCell;
        var isMerged2 = cell2.IsMergedCell;

        // Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        var internalCell1 = (TableCell)cell1;
        var internalCell2 = (TableCell)cell2;
        internalCell1.RowIndex.Should().Be(internalCell2.RowIndex);
        internalCell1.ColumnIndex.Should().Be(internalCell2.ColumnIndex);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_true_When_0x1_and_1x1_cells_are_merged()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slide(4).Shape(4).Table;
        var cell1 = table[0, 1];
        var cell2 = table[1, 1];

        // Act
        var isMerged1 = cell1.IsMergedCell;
        var isMerged2 = cell2.IsMergedCell;

        // Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        var internalCell1 = (TableCell)cell1;
        var internalCell2 = (TableCell)cell2;
        internalCell1.RowIndex.Should().Be(internalCell2.RowIndex);
        internalCell1.ColumnIndex.Should().Be(internalCell2.ColumnIndex);
    }

    [Test]
    [TestCase(0, 0, 0, 1)]
    [TestCase(0, 1, 0, 0)]
    public void MergeCells_MergesSpecifiedCellsRange(int rowIdx1, int colIdx1, int rowIdx2, int colIdx2)
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 4);
        var mStream = new MemoryStream();

        // Act
        table.MergeCells(table[rowIdx1, colIdx1], table[rowIdx2, colIdx2]);

        // Assert
        table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
        table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();

        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 4);
        table[rowIdx1, colIdx1].IsMergedCell.Should().BeTrue();
        table[rowIdx2, colIdx2].IsMergedCell.Should().BeTrue();
    }

    [Test]
    public void AltText_Setter_sets_alternative_text()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Shape("Table 1");

        // Act
        table.AltText = "Alt text";

        // Assert
        table.AltText.Should().Be("Alt text");
        pres.Validate();
    }

    [Test]
    [TestCase(true, false, false, false, false, false)]
    [TestCase(false, true, false, false, false, false)]
    [TestCase(false, false, true, false, false, false)]
    [TestCase(false, false, false, true, false, false)]
    [TestCase(false, false, false, false, true, false)]
    [TestCase(false, false, false, false, false, true)]
    public void TableStyleOptions_property_setters_set_table_style_options(
        bool hasHeaderRow,
        bool hasTotalRow,
        bool hasBandedRows,
        bool hasFirstColumn,
        bool hasLastColumn,
        bool hasBandedColumns)
    {
        // Arrange
        var mStream = new MemoryStream();
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 3, 2);
        var table = slide.Shapes.Last().Table;

        // Act
        table.StyleOptions.HasHeaderRow = hasHeaderRow;
        table.StyleOptions.HasTotalRow = hasTotalRow;
        table.StyleOptions.HasBandedRows = hasBandedRows;
        table.StyleOptions.HasFirstColumn = hasFirstColumn;
        table.StyleOptions.HasLastColumn = hasLastColumn;
        table.StyleOptions.HasBandedColumns = hasBandedColumns;

        // Assert
        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = pres.Slide(1).Shapes.Last().Table;
        table.StyleOptions.HasHeaderRow.Should().Be(hasHeaderRow);
        table.StyleOptions.HasTotalRow.Should().Be(hasTotalRow);
        table.StyleOptions.HasBandedRows.Should().Be(hasBandedRows);
        table.StyleOptions.HasFirstColumn.Should().Be(hasFirstColumn);
        table.StyleOptions.HasLastColumn.Should().Be(hasLastColumn);
        table.StyleOptions.HasBandedColumns.Should().Be(hasBandedColumns);
    }

    [Test]
    public void TableStyleOptions_property_getters_return_table_style_options()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 3, 2);
        var table = slide.Shapes.Last().Table;

        // Act
        var options = table.StyleOptions;

        // Assert
        options.HasHeaderRow.Should().BeTrue();
        options.HasTotalRow.Should().BeFalse();
        options.HasBandedRows.Should().BeTrue();
        options.HasFirstColumn.Should().BeFalse();
        options.HasLastColumn.Should().BeFalse();
        options.HasBandedColumns.Should().BeFalse();
    }

    [Test]
    public void Height_Setter_should_proportionally_increase_the_row_heights_When_the_new_table_height_is_bigger()
    {
        // Arrange
        var pres = new Presentation(p=>p.Slide());
        pres.Slide(1).Shapes.AddTable(10, 10, 2, 2);
        var tableShape = pres.Slide(1).Shapes.Last();
        var table = tableShape.Table;

        // Act
        tableShape.Height *= 1.5m;

        // Assert
        table.Rows[0].Height.Should().BeApproximately(43, 0.01m);
        table.Rows[1].Height.Should().Be(43);
        pres.Validate();
    }

    [Test]
    public void Rows_Add_adds_row_at_the_specified_index()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;
        var rowsCountBefore = table.Rows.Count;

        // Act
        table.Rows.Add(1);

        // Assert
        table.Rows.Should().HaveCount(rowsCountBefore + 1);
        table.Rows[1].Cells[0].TextBox.Text.Should().BeEmpty();
        pres = SaveAndOpenPresentation(pres);
        table = pres.Slide(1).Shape("Table 1").Table;
        table.Rows.Should().HaveCount(rowsCountBefore + 1);
        pres.Validate();
    }

    [Test]
    public void Rows_Add_adds_a_new_row_at_the_specified_index_with_the_template_height()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;
        var templateRowIndex = 0;
        var templateRowHeight = table.Rows[templateRowIndex].Height;

        // Act
        table.Rows.Add(1, templateRowIndex);

        // Assert
        pres = SaveAndOpenPresentation(pres);
        table = pres.Slide(1).Shape("Table 1").Table;
        table.Rows[1].Height.Should().Be(templateRowHeight);
        pres.Validate();
    }

    [Test]
    public void Rows_Add_adds_a_new_row_with_template_font_color()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Shape("Table 1").Table;
        var templateRowIndex = 0;
        var templateFontColor = table.Rows[templateRowIndex].Cells[0].TextBox.Paragraphs[0].Portions[0].Font!.Color.Hex;

        // Act
        table.Rows.Add(1, templateRowIndex);

        // Assert
        pres = SaveAndOpenPresentation(pres);
        pres.Slide(1).Shape("Table 1").Table.Rows[1].Cells[0].Fill.Color.Should().Be(templateFontColor);
    }
}