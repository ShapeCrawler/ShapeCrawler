using System.Reflection;

namespace ShapeCrawler.DevTests.Helpers;

internal class JSONPresentation
{
    public JSONSlide[] Slides;
    
    internal Presentation AsSCPresentation()
    {
        var scPres = new Presentation();
        var firstSlide = this.Slides[0];
        foreach (var shape in firstSlide.Shapes)
        {
            if (shape.VideoContent is not null)
            {
                var videoStream = Assembly.GetExecutingAssembly().GetResourceStream(shape.VideoContent);
                var scFirstSlide = scPres.Slide(1);
                scFirstSlide.Shapes.AddVideo(10,10, videoStream);
                scFirstSlide.Shapes.Last().Name = shape.Name;
            }
            else if (shape.TextContent is not null)
            {
                var scFirstSlide = scPres.Slide(1);
                scFirstSlide.Shapes.AddShape(10,10,10,10, Geometry.Rectangle, shape.TextContent);
                scFirstSlide.Shapes.Last().Name = shape.Name;
            }
        }
        
        return scPres;
    }
}