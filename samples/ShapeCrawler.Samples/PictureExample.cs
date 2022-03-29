using ShapeCrawler;

internal class PictureExample
{
    internal static void ReadPicture()
    {
        using var presentation = SCPresentation.Open(@"test.pptx", true);
        var slide = presentation.Slides[0];
        
        // Get picture shape by name
        var pictureShape = slide.Shapes.GetByName<IPicture>("Picture 1");

        // Get MIME type of image, eg. "image/png"
        var mimeType = pictureShape.Image.MIME;
    }
}