namespace ShapeCrawler.Examples;

public class Font
{
    [Test, Explicit]
    public void Set_Latin_font()
    {
        var pres = new Presentation();
        var slide = pres.Slide(1);
        slide.Shapes.AddShape(0, 0, 100, 100, Geometry.Rectangle, "Test");
        var addedShape = slide.Shapes.Last();
        var font = addedShape.TextBox!.Paragraphs[0].Portions[0].Font!;
        
        font.LatinName = "Times New Roman";
    }
}