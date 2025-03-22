using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml.Spreadsheet;
using FluentAssertions;
using NUnit.Framework;
using ShapeCrawler.Tables;
using ShapeCrawler.Tests.Unit.Helpers;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Tests.Unit;

public class TableTests : SCTest
{
	[Test]
	public void TableStyle_Getter_return_style_of_table()
	{
		// Arrange
		var pres = new Presentation(TestAsset("009_table.pptx"));
        var table = pres.Slide(3).Table("Таблица 4");

		// Act 
		var tableStyle = table.TableStyle;

		// Assert
		tableStyle.Should().BeEquivalentTo(TableStyle.MediumStyle2Accent1);
		pres.Validate();
	}

	[Test]
	public void TableStyle_Setter_sets_style()
	{
		// Arrange
		var pres = new Presentation(TestAsset("009_table.pptx"));
		var table = pres.Slide(3).Table("Таблица 4");
		var mStream = new MemoryStream();

		// Act
		table.TableStyle = TableStyle.ThemedStyle1Accent4;

		// Assert
		table.TableStyle.Should().BeEquivalentTo(TableStyle.ThemedStyle1Accent4);
		pres.Save(mStream);
		pres = new Presentation(mStream);
        table = pres.Slide(3).Table("Таблица 4");
		table.TableStyle.Should().BeEquivalentTo(TableStyle.ThemedStyle1Accent4);
		pres.Validate();
	}

	[Test]
	public void CellBorder_Getter_return_bottom_border_color()
	{
		// Arrange
		var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Table("Table 1");

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
		var table = pres.Slide(1).Table("Table 1");
		var mStream = new MemoryStream();

        // Act
        table[1, 2].RightBorder.Color = "00FF00";

		// Assert
		table[1, 2].RightBorder.Color.Should().Be("00FF00");
		pres.Save(mStream);

		pres = new Presentation(mStream);
        table = pres.Slide(1).Table("Table 1");
		table[1, 2].RightBorder.Color.Should().Be("00FF00");
		pres.Validate();
	}

    [Test]
	public void CellMargins_Getter_return_cell_margin()
	{
		// Arrange
		var pres = new Presentation(TestAsset("table-case004.pptx"));
        var table = pres.Slide(1).Table("Table 2");
        
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
		var table = pres.Slide(1).Table("Table 2");
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
        table = pres.Slide(1).Table("Table 2");
		
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
        var table = pres.Slide(1).Table("Table 2");
        
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
		var table = pres.Slide(1).Table("Table 2");
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
        table = pres.Slide(1).Table("Table 2");
		
        tableCell = table[1, 1];
		tableCell.TextBox.VerticalAlignment.Should().Be(TextVerticalAlignment.Middle);
        tableCell.TextBox.Paragraphs[0].HorizontalAlignment.Should().Be(TextHorizontalAlignment.Center);
		
		pres.Validate();
	}

	[Test]
    public void RemoveColumnAt_removes_column_by_specified_index()
    {
        // Arrange
        var ms = new MemoryStream();
        var pptx = TestAsset("table-case001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var expectedColumnsCount = table.Columns.Count - 1;

        // Act
        table.Columns.RemoveAt(1);

        // Assert
        table.Columns.Should().HaveCount(expectedColumnsCount);
        pres.Save(ms);
        pres = new Presentation(ms);
        table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        table.Columns.Should().HaveCount(expectedColumnsCount);
        pres.Validate();
    }
    
    [Test]
    public void AddColumn_adds_column()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Table("Table 1");
        var expectedColumnsCount = table.Columns.Count + 1;

        // Act
        table.Columns.Add();

        // Assert
        table.Columns.Should().HaveCount(expectedColumnsCount);
        pres.Validate();
    }
    
    [Test]
    public void Columns_Add_sets_width_of_all_columns_proportionally()
    {
        // Arrange
        var pptx = TestAsset("table-case003.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var columnsCountBefore = table.Columns.Count;
        var columnWidthBefore = table.Columns.Select(c => c.Width).ToList();
        var totalWidthBefore = table.Columns.Sum(c => c.Width);
        var newTotalWidth = totalWidthBefore + table.Columns[columnsCountBefore - 1].Width;
        
        // Act
        table.Columns.Add();

        // Assert
        var widthRatio = (double)totalWidthBefore / newTotalWidth;
        table.Columns.Select(c => c.Width).ToList().Take(columnsCountBefore).Should()
            .BeEquivalentTo(columnWidthBefore.Select(w => (int)(w * widthRatio)));
    }
    
    [Test]
    public void InsertColumnAfter_inserts_column_after_the_specified_column_number()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slide(1).Table("Table 1");

        // Act
        table.Columns.InsertAfter(1);
        var cell = table.Cell(1, 2);

        // Assert
        cell.TextBox.Text.Should().BeEmpty("because before adding column the cell (1,2) was not empty.");
        pres.Validate();
    }
    
    [Test]
    public void Columns_InsertAfter_sets_width_of_all_columns_proportionally()
    {
        // Arrange
        var pptx = TestAsset("table-case003.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var columnsCountBefore = table.Columns.Count;
        var columnWidthBefore = table.Columns.Select(c => c.Width).ToList();
        var totalWidthBefore = table.Columns.Sum(c => c.Width);
        var newTotalWidth = totalWidthBefore + table.Columns[2].Width;
        
        // Act
        table.Columns.InsertAfter(3);

        // Assert
        var widthRatio = (double)totalWidthBefore / newTotalWidth;
        table.Columns.Select(c => c.Width).ToList().Take(columnsCountBefore).Should()
            .BeEquivalentTo(columnWidthBefore.Select(w => (int)(w * widthRatio)));
    }
    
    [Test]
    public void Columns_Duplicate_increases_column_count_by_one()
    {
        // Arrange
        var pptx = TestAsset("table-case001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var column = table.Columns[0];
        var columnsCountBefore = table.Columns.Count;

        // Act
        column.Duplicate();
        
        // Assert
        table.Columns.Should().HaveCount(columnsCountBefore + 1);
        pres.Validate();
    }
    
    [Test]
    public void Columns_Duplicate_copies_column_with_all_its_cells()
    {
        // Arrange
        var pptx = TestAsset("table-case002.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var column = table.Columns[0];

        // Act
        column.Duplicate();
        
        // Assert
        foreach (var row in table.Rows)
        {
            row.Cells.Should().HaveCount(table.Columns.Count);
            row.Cells[0].TextBox.Text.Should().Be(row.Cells[table.Columns.Count - 1].TextBox.Text);
        }
    }
    
    [Test]
    public void Columns_Duplicate_copies_middle_column_with_all_its_cells()
    {
        // Arrange
        var pptx = TestAsset("table-case003.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var column = table.Columns[1];

        // Act
        column.Duplicate();
        
        // Assert
        foreach (var row in table.Rows)
        {
            row.Cells.Should().HaveCount(table.Columns.Count);
            row.Cells[1].TextBox.Text.Should().Be(row.Cells[table.Columns.Count - 1].TextBox.Text);
        }
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
    public void Columns_Duplicate_sets_width_of_all_columns_proportionally()
    {
        // Arrange
        var pptx = TestAsset("table-case003.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");
        var column = table.Columns[0];
        var columWidthBefore = table.Columns.Select(c => c.Width).ToList();
        var totalWidthBefore = table.Columns.Sum(c => c.Width);
        var newTotalWidth = totalWidthBefore + column.Width;
        
        // Act
        column.Duplicate();

        // Assert
        var widthRatio = (double)totalWidthBefore / newTotalWidth;
        table.Columns.Select(c => c.Width).ToList().Take(columWidthBefore.Count).Should()
            .BeEquivalentTo(columWidthBefore.Select(w => (int)(w * widthRatio)));
    }
    
    [Test]
    public void Rows_Add_adds_row()
    {
        // Arrange
        var pptx = TestAsset("table-case001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[0].Shapes.GetByName<ITable>("Table 1");

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
    public void Row_Height_Getter_returns_row_height_in_points()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
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
        ITable table = (ITable)new Presentation(TestAsset("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);

        // Act
        int columnsCount = table.Columns.Count;

        // Assert
        columnsCount.Should().Be(3);
    }

    [Test]
    public void Column_Width_Getter_returns_width_of_column_in_pixels()
    {
        // Arrange
        var table = (ITable)new Presentation(TestAsset("001.pptx")).Slides[1].Shapes.First(sp => sp.Id == 4);

        // Act
        var columnWidth = table.Columns[0].Width;

        // Assert
        columnWidth.Should().Be(367);
    }

    [Test]
    public void Column_Width_Setter_sets_width_of_column()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 3);
        const int newColumnWidth = 427;
        var mStream = new MemoryStream();

        // Act
        table.Columns[0].Width = newColumnWidth;

        // Assert
        table.Columns[0].Width.Should().Be(newColumnWidth);

        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = (ITable)pres.Slides[1].Shapes.First(sp => sp.Id == 3);
        table.Columns[0].Width.Should().Be(newColumnWidth);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_True_When_cell_is_merged()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var row = pres.Slides[1].Shapes.GetByName<ITable>("Table 4").Rows[1];
        var cell1X0 = row.Cells[0];
        var cell1X1 = row.Cells[1];

        // Act
        var isMerged1 = cell1X0.IsMergedCell;
        var isMerged2 = cell1X1.IsMergedCell;

        // Act-Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        
        var cell1 = (TableCell) cell1X0;
        var cell2 = (TableCell)cell1X1;
        cell1.RowIndex.Should().Be(cell2.RowIndex);
        cell1.ColumnIndex.Should().Be(cell2.ColumnIndex);
    }
    
    [Test]
    public void Row_Cell_TopBorder_Width_Setter_sets_top_border_width_in_points()
    {
        // Arrange
        var pres = new Presentation();
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
        var cell = pres.Slides[0].Table("Table 1")[0, 0];
        
        // Act
        cell.LeftBorder.Width = 2;  
        
        // Assert
        cell.LeftBorder.Width.Should().Be(2);
    }
    
    [Test]
    public void Row_Cell_TopBorder_Width_Getter_returns_top_border_width_in_points()
    {
        // Arrange
        var pres = new Presentation();
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
    public void Row_Height_Setter_sets_height_of_the_table_row_in_points()
    {
        // Arrange
        var pres = new Presentation(TestAsset("table-case001.pptx"));
        var table = pres.Slides[0].Table("Table 1");
        var row = table.Rows[0];

        // Act
        row.Height = 39;

        // Assert
        row.Height.Should().Be(39);
        table.Height.Should().BeApproximately(39.2m, 0.01m);
    }

    [Test(Description = "MergeCells #1")]
    public void MergeCells_Merges0x0And0x1CellsOf2x2Table()
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

    [Test(Description = "MergeCells #2")]
    public void MergeCells_Merges0x1And0x2CellsOf3x2Table()
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

    [Test(Description = "MergeCells #3")]
    public void MergeCells_Merges0x0And0x1And0x2CellsOf3x2Table()
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

    [Test(Description = "MergeCells #4")]
    public void MergeCells_Merges0x0And0x1MergedCellsWith0x2CellIn3x2Table()
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

    [Test(Description = "MergeCells #5")]
    public void MergeCells_merges_0x0_and_1x0_cells_of_2x2_table()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[2].Shapes.GetById<ITable>(5);
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
        var pres = new Presentation();
        pres.Slides[0].Shapes.AddTable(10, 10, 3, 4);
        var table = (ITable)pres.Slides[0].Shapes.Last();
        table[1, 0].TextBox.Text = "A";
        table[3, 0].TextBox.Text = "B";

        // Act
        table.MergeCells(table[1, 0], table[2, 0]);

        // Assert
        table[1, 0].TextBox.Text.Should().Be("A");
    }

    [Test(Description = "MergeCells #6")]
    public void MergeCells_Merges0x1And1x1CellsOf3x2Table()
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

    [Test(Description = "MergeCells #7")]
    public void MergeCells_Merges0x0To1x1RangeOf3x3Table()
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
    public void MergeCells_MergesMergedCellWithNonMergedCell()
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
        
        var cell1 = (TableCell) table[1, 1];
        var cell2 = (TableCell)table[1, 2];
        cell1.RowIndex.Should().Be(cell2.RowIndex);
        cell1.ColumnIndex.Should().Be(cell2.ColumnIndex);
        
        table[3, 2].IsMergedCell.Should().BeFalse();

        presentation.Save(mStream);
        presentation = new Presentation(mStream);
        table = (ITable)presentation.Slides[1].Shapes.First(sp => sp.Id == 5);
        table[1, 1].IsMergedCell.Should().BeTrue();
        table[1, 2].IsMergedCell.Should().BeTrue();
        
        cell1 = (TableCell) table[1, 1];
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
        var table = pres.Slides[3].Table("Table 2");
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

    [Test(Description = "MergeCells #8")]
    public void MergeCells_converts_2X1_table_into_1X1_when_all_cells_are_merged()
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

    [Test(Description = "MergeCells #9")]
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

    [Test(Description = "MergeCells #10")]
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

        void AssertTable(ITable tableSc, int expectedMergedColumnWidth)
        {
            tableSc.Columns.Should().HaveCount(2);
            tableSc.Columns[0].Width.Should().Be(expectedMergedColumnWidth);
            tableSc.Rows.Should().HaveCount(1);
            tableSc.Rows[0].Cells.Should().HaveCount(2);
        }
    }

    [Test(Description = "MergeCells #11")]
    public void MergeCells_merges_0x1_and_0x2_cells()
    {
        // Arrange
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[3].Shapes.GetById<ITable>(6);
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
        var pptx = TestAsset("001.pptx");
        var pres = new Presentation(pptx);
        var table = pres.Slides[3].Shapes.GetById<ITable>(6);

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
        var pres = new Presentation();
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
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 4, 2);
        var table = slide.Shapes.Last() as ITable;
        
        // Act
        table.MergeCells(table[0, 0], table[0, 1]);
        table.MergeCells(table[1, 1], table[1, 2]);
        
        // Assert
        var cell1 = (TableCell) table[1, 1];
        var cell2 = (TableCell)table[1, 2];
        cell1.RowIndex.Should().Be(cell2.RowIndex);
        cell1.ColumnIndex.Should().Be(cell2.ColumnIndex);
    }

    [Test]
    public void MergeCells_merges_0x1_and_1x1()
    {
        // Arrange
        var pres = new Presentation();
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
        var table = pres.Slides[1].Shapes.GetById<ITable>(3);
        var cell1 = table[0, 0];
        var cell2 = table[1, 0];

        // Act
        var isMerged1 = cell1.IsMergedCell;
        var isMerged2 = cell2.IsMergedCell;

        // Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        var internalCell1 = (TableCell) cell1;
        var internalCell2 = (TableCell) cell2;
        internalCell1.RowIndex.Should().Be(internalCell2.RowIndex);
        internalCell1.ColumnIndex.Should().Be(internalCell2.ColumnIndex);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_true_When_1x1_and_2x1_cells_are_merged()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slides[1].Shapes.GetByName<ITable>("Table 5");
        var cell1 = table[1, 1];
        var cell2 = table[2, 1];

        // Act
        var isMerged1 = cell1.IsMergedCell;
        var isMerged2 = cell2.IsMergedCell;

        // Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        var internalCell1 = (TableCell) cell1;
        var internalCell2 = (TableCell) cell2;
        internalCell1.RowIndex.Should().Be(internalCell2.RowIndex);
        internalCell1.ColumnIndex.Should().Be(internalCell2.ColumnIndex);
    }

    [Test]
    public void Row_Cell_IsMergedCell_returns_true_When_0x1_and_1x1_cells_are_merged()
    {
        // Arrange
        var pres = new Presentation(TestAsset("001.pptx"));
        var table = pres.Slides[3].Shapes.GetById<ITable>(4);
        var cell1 = table[0, 1];
        var cell2 = table[1, 1];

        // Act
        var isMerged1 = cell1.IsMergedCell;
        var isMerged2 = cell2.IsMergedCell;

        // Assert
        isMerged1.Should().BeTrue();
        isMerged2.Should().BeTrue();
        var internalCell1 = (TableCell) cell1;
        var internalCell2 = (TableCell) cell2;
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
        var table = pres.Slide(1).Table("Table 1");

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
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 3, 2);
        var table = slide.Shapes.Last<ITable>();

        // Act
        table.TableStyleOptions.HasHeaderRow = hasHeaderRow;
        table.TableStyleOptions.HasTotalRow = hasTotalRow;
        table.TableStyleOptions.HasBandedRows = hasBandedRows;
        table.TableStyleOptions.HasFirstColumn = hasFirstColumn;
        table.TableStyleOptions.HasLastColumn = hasLastColumn;
        table.TableStyleOptions.HasBandedColumns = hasBandedColumns;

        // Assert
        pres.Save(mStream);
        pres = new Presentation(mStream);
        table = pres.Slides[0].Shapes.Last<ITable>();
        table.TableStyleOptions.HasHeaderRow.Should().Be(hasHeaderRow);
        table.TableStyleOptions.HasTotalRow.Should().Be(hasTotalRow);
        table.TableStyleOptions.HasBandedRows.Should().Be(hasBandedRows);
        table.TableStyleOptions.HasFirstColumn.Should().Be(hasFirstColumn);
        table.TableStyleOptions.HasLastColumn.Should().Be(hasLastColumn);
        table.TableStyleOptions.HasBandedColumns.Should().Be(hasBandedColumns);
    }
    
    [Test]
    public void TableStyleOptions_property_getters_return_table_style_options()
    {
        // Arrange
        var pres = new Presentation();
        var slide = pres.Slides[0];
        slide.Shapes.AddTable(0, 0, 3, 2);
        var table = slide.Shapes.Last<ITable>();

        // Act
        var options = table.TableStyleOptions;

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
        var pres = new Presentation();
        pres.Slide(1).Shapes.AddTable(10, 10, 2, 2);
        var addedTable = pres.Slide(1).Shapes.Last<ITable>();
        
        // Act
        var currentRowsHeight = addedTable.Rows[0].Height;
        addedTable.Height *= 1.5m;
        
        // Assert
        addedTable.Rows[0].Height.Should().Be((int)(currentRowsHeight * 1.5));
        addedTable.Rows[1].Height.Should().Be((int)(currentRowsHeight * 1.5));
        pres.Validate();
    }
}