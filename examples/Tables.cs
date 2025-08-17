namespace ShapeCrawler.Examples;

public class Tables
{
    [Test, Explicit]
    public void Create_table()
    {
        using var pres = new Presentation("pres.pptx");
        var shapeCollection = pres.Slide(1).Shapes;

        shapeCollection.AddTable(x: 50, y: 100, columnsCount: 3, rowsCount: 2);
        var addedTable = (ITable)shapeCollection.Last();
        var cell = addedTable[0, 0];
        cell.TextBox.SetText("Hi, Table!");

        pres.Save();
    }

    [Test, Explicit]
    public void Get_table_properties()
    {
        using var pres = new Presentation("helloWorld.pptx");
        var slide = pres.Slide(1);
        
        var table = (ITable)slide.Shapes.First(sp => sp is ITable);
        
        var rowsCount = table.Rows.Count;
        
        var cellsCount = table.Rows[0].Cells.Count;

        // Print message if the cell belongs to merged cells group
        foreach (var row in table.Rows)
        {
            foreach (var cellItem in row.Cells)
            {
                if (cellItem.IsMergedCell)
                {
                    Console.WriteLine("The cell is a part of a merged cells group.");
                }
            }
        }

        // Get column width
        var column = table.Columns[0];
        var columnWidth = column.Width;

        // Get row height
        var rowHeight = table.Rows[0].Height;

        // Get cell with row index 0 column index 1
        var cell = table[0, 1];
    }

    public static void Merge_cells()
    {
        using var pres = new Presentation("pres.pptx");
        var slide = pres.Slide(1);
        var table = (ITable)slide.Shapes.First(sp => sp is ITable);
        
        table.MergeCells(table[0, 0], table[0, 1]);
    }

    public static void Remove_row()
    {
        using var pres = new Presentation("table.pptx");
        var slide = pres.Slide(1);
        var table = slide.Shapes.Shape("Table 1").Table;

        table.Rows.RemoveAt(0);
    }
    
    public static void Add_row()
    {
        using var pres = new Presentation("presentation.pptx");
        var slide = pres.Slide(1);
        var table = slide.Shapes.Shape("Table 1").Table;

        // Add a new row at the 1 index using row with the index 0 as a template
        table.Rows.Add(1,0);
    }
}