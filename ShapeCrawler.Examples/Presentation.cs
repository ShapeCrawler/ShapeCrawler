namespace ShapeCrawler.Examples;

public class Presentation
{
    [Test]
    [Explicit]
    public void Get_markdown()
    {
        var pres = new ShapeCrawler.Presentation("pres.pptx");

        var presMarkdown = pres.AsMarkdown();
    }
}