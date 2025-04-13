namespace ShapeCrawler.Examples;

public class Slide
{
    [Test]
    [Explicit]
    public void Remove_slide()
    {
        var pres = new Presentation("pres.pptx");
        
        pres.Slide(1).Remove();
        
        pres.Save();
    }
}