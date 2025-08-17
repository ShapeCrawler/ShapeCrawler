namespace ShapeCrawler.Examples;

public class ShapeCollection
{
    [Test]
    public void Groups_shapes()
    {
        using var pres = new Presentation("pres.pptx");
        var shapes = pres.Slide(1).Shapes;
        var shape1 = shapes.Shape("Shape 1");
        var shape2 = shapes.Shape("Shape 2");
        
        var group = shapes.Group([shape1, shape2]);
    }
}