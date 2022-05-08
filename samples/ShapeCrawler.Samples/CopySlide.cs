using ShapeCrawler;

internal class CopySlide
{
    internal void CopySlideSample()
    {
        using var sourcePre = SCPresentation.Open(@"source.pptx", false);
        using var destPre = SCPresentation.Open(@"dest.pptx", true);

        var copyingSlide = sourcePre.Slides[0];
        destPre.Slides.Add(copyingSlide);
    }
}