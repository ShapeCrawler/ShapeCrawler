using System.IO;
using System.Linq;
using ShapeCrawler.Models.TextShape;

namespace ShapeCrawler.Tests.Unit
{
    public class TestHelper
    {
        public static Paragraph GetParagraph(PresentationSc presentation, ElementRequest paragraphRequest)
        {
            return presentation.Slides[paragraphRequest.SlideIndex]
                                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId)
                                .TextFrame.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static Paragraph GetParagraph(MemoryStream presentationStream, ElementRequest paragraphRequest)
        {
            PresentationSc presentation = PresentationSc.Open(presentationStream, false);

            return presentation.Slides[paragraphRequest.SlideIndex]
                                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId)
                                .TextFrame.Paragraphs[paragraphRequest.ParagraphIndex];
        }
    }
}