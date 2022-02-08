namespace ShapeCrawler.Samples;

internal class TextExample
{
    internal void ReadAutoShape()
    {
        using var presentation = SCPresentation.Open(@"test.pptx", true);
        var autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1");
        var paragraph = autoShape.TextBox.Paragraphs[0];

        // Read alignment
        var alignment = paragraph.Alignment;

        // Update alignment
        alignment = TextAlignment.Center;
    }
}