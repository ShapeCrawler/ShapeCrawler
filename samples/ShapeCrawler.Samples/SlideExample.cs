using ShapeCrawler;

internal class SlideExample
{
    internal static void ReadSlide()
    {
        using var presentation = SCPresentation.Open(@"test.pptx", true);
        var slide = presentation.Slides[0];
        
        // Get background image byte content
        var bytes = slide.Background.GetBytes();
    }
}