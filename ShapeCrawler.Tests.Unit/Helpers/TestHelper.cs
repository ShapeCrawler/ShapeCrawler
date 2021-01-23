using ShapeCrawler.Texts;
using System.IO;
using System.Linq;

namespace ShapeCrawler.Tests.Unit.Helpers
{
    public class TestHelper
    {
        public static ParagraphSc GetParagraph(PresentationSc presentation, ElementRequest paragraphRequest)
        {
            return presentation.Slides[paragraphRequest.SlideIndex]
                                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId)
                                .TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static ParagraphSc GetParagraph(MemoryStream presentationStream, ElementRequest paragraphRequest)
        {
            PresentationSc presentation = PresentationSc.Open(presentationStream, false);

            return presentation.Slides[paragraphRequest.SlideIndex]
                                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId)
                                .TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }
    }
}