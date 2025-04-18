namespace ShapeCrawler.Examples;

public class Slides
{
    [Test, Explicit]
    public void Remove_slide()
    {
        // Remove first slide
        using var pres = new Presentation("test.pptx");
        var removingSlide = pres.Slides.First();
        removingSlide.Remove();

        // Move second slide to the first position
        pres.Slides[1].Number = 1;

        // Copy slide to another presentation
        var sourcePresentation = new Presentation("source.pptx");
        var targetPres = new Presentation("dest.pptx");
        var copyingSlide = sourcePresentation.Slides[1];
        targetPres.Slides.Add(copyingSlide);
        
        pres.Save();
    }
}