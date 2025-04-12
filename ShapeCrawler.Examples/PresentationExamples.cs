namespace ShapeCrawler.Examples;

public class PresentationExamples
{
    [Test]
    [Explicit]
    public void Get_markdown()
    {
        var pres = new Presentation("pres.pptx");

        var presMarkdown = pres.AsMarkdown();
    }
}