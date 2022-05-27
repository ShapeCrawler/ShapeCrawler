using ShapeCrawler;

namespace TextExample;

internal class TextSample
{
    internal void ReadAndUpdateAutoShape()
    {
        using var presentation = SCPresentation.Open(@"test.pptx", true);
        var autoShape = presentation.Slides[0].Shapes.GetByName<IAutoShape>("TextBox 1");
        var paragraph = autoShape.TextBox.Paragraphs[0];

        // Read alignment
        var alignment = paragraph.Alignment;

        // Update alignment
        paragraph.Alignment = TextAlignment.Center;
        
        // Add/Update Hyperlink
        paragraph.Portions[0].Hyperlink = "https://github.com/ShapeCrawler/ShapeCrawler";
    }
}