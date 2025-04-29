namespace ShapeCrawler.Examples;

public class Slides
{
    [Test, Explicit]
    public void Remove_slide()
    {
        using var pres = new Presentation("presentation.pptx");
        
        pres.Slide(1).Remove();
    }
    
    [Test, Explicit]
    public void Update_slide_number()
    {
        using var pres = new Presentation("presentation.pptx");
        
        pres.Slide(2).Number = 1;
    }
    
    [Test, Explicit]
    public void Copy_slide_to_another_presentation()
    {
        using var sourcePres = new Presentation("source.pptx");
        using var targetPres = new Presentation("target.pptx");
        var copyingSlide = sourcePres.Slides[1];
        
        targetPres.Slides.Add(copyingSlide);
    }
}