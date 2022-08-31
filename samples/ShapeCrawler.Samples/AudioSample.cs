using ShapeCrawler;

internal class AudioSample
{
    internal void Home_Examples_Audio()
    {
        var pres = SCPresentation.Open("test.pptx", true);
        var audioShape = pres.Slides[0].Shapes.GetByName<IAudioShape>("Audio 1");
        var audioBytes = audioShape.BinaryData;
    }
}