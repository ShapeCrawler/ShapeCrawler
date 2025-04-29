namespace ShapeCrawler.Examples;

public class Tables
{
    [Test, Explicit]
    public void Create_table()
    {
        using var pres = new Presentation("pres.pptx");
        var shapeCollection = pres.Slides[0].Shapes;

        shapeCollection.AddTable(x: 50, y: 100, columnsCount: 3, rowsCount: 2);
        var addedTable = (ITable)shapeCollection.Last();
        var cell = addedTable[0, 0];
        cell.TextBox.SetText("Hi, Table!");

        pres.Save();
    }
}