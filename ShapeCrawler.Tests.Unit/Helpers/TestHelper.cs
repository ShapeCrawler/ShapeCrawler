using ShapeCrawler.Texts;
using System.IO;
using System.Linq;
using ShapeCrawler.AutoShapes;
using ShapeCrawler.SlideMaster;

namespace ShapeCrawler.Tests.Unit.Helpers
{
    public class TestHelper
    {
        public static ParagraphSc GetParagraph(PresentationSc presentation, ElementRequest paragraphRequest)
        {
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static ParagraphSc GetParagraph(MemoryStream presentationStream, ElementRequest paragraphRequest)
        {
            PresentationSc presentation = PresentationSc.Open(presentationStream, false);
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }
    }
}