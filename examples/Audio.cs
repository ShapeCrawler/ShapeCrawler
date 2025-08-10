namespace ShapeCrawler.Examples;

public class Audio
{
    [Test, Explicit]
    public static void Add_Audio_shape()
    {
        using var pres = new Presentation("audio.pptx");
        var shapes = pres.Slide(1).Shapes;
        using var audioStream = File.OpenRead("audio.mp3");
        shapes.AddAudio(x: 300, y: 100, audioStream);
        var addedAudio = shapes.Last().Media;
        pres.Save();

        // Get byte content
        var audioBytes = addedAudio.AsByteArray();
        
        // Get MIME type, e.g. 'audio/mpeg'
        var mime = addedAudio.MIME;
    }
}