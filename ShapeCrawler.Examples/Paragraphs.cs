namespace ShapeCrawler.Examples;

public class Tests
{
    [Test]
    [Explicit]
    public void Add_paragraph()
    {
        var pres = new Presentation();
        var shapes = pres.Slides[0].Shapes;
        shapes.AddShape(100, 100, 200, 200);
        var addedShape = shapes.Last();
        addedShape.TextBox!.Text = "Hello World!";
        
        addedShape.TextBox.Paragraphs.Add();
        var addedParagraph = addedShape.TextBox.Paragraphs.Last();
        addedParagraph.Text = "I'm ShapeCrawler";
        
        pres.Save(@"C:\temp\pres.pptx");
    }
}