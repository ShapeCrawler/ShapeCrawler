namespace ShapeCrawler.Examples;

public class ShapeExamples
{
    [Test, Explicit]
    public void Set_shape_fill()
    {
        using var pres = new Presentation("pres.pptx");
        var shape = pres.Slide(1).Shapes.Shape<IShape>("AutoShape 1");
        const string green = "00FF00";

        shape.Fill!.SetColor(green);
    }

    [Test, Explicit]
    public void Update_shape_video_content()
    {
        using var pres = new Presentation("pres.pptx");
        var videoShape = pres.Slide(1).Shape("Video");
        using var videoContent = File.OpenRead("video.mp4");
        
        videoShape.SetVideo(videoContent);
    }
}