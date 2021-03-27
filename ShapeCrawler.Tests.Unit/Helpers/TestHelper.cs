using System.IO;
using System.Linq;
using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Tests.Unit.Helpers
{
    public class TestHelper
    {
        public static SCParagraph GetParagraph(SCPresentation presentation, ElementRequest paragraphRequest)
        {
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static SCParagraph GetParagraph(MemoryStream presentationStream, ElementRequest paragraphRequest)
        {
            SCPresentation presentation = SCPresentation.Open(presentationStream, false);
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static Portion GetPortion(SCPresentation presentation, ElementRequest elementRequest)
        {
            IAutoShape autoShape = (IAutoShape)presentation.Slides[elementRequest.SlideIndex].Shapes.First(sp => sp.Id == elementRequest.ShapeId);
            
            return autoShape.TextBox.Paragraphs[elementRequest.ParagraphIndex].Portions[elementRequest.PortionIndex];
        }
    }
}