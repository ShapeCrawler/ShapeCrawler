namespace ShapeCrawler.Examples;

public class Shape
{
    [Test, Explicit]
    public void Set_shape_fill()
    {
        using var pres = new Presentation("pres.pptx");
        var shape = pres.Slide(1).Shapes.Shape<IShape>("AutoShape 1");
        const string green = "00FF00";

        shape.Fill!.SetColor(green);
    }
}