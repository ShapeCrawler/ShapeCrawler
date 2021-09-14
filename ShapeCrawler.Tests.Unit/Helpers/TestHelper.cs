using System.IO;
using System.Linq;
using ShapeCrawler.AutoShapes;

namespace ShapeCrawler.Tests.Unit.Helpers
{
    public static class TestHelper
    {
        public static IParagraph GetParagraph(SCPresentation presentation, SlideElementQuery paragraphRequest)
        {
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static IParagraph GetParagraph(MemoryStream presentationStream, SlideElementQuery paragraphRequest)
        {
            SCPresentation presentation = SCPresentation.Open(presentationStream, false);
            IAutoShape autoShape = presentation.Slides[paragraphRequest.SlideIndex]
                .Shapes.First(sp => sp.Id == paragraphRequest.ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphRequest.ParagraphIndex];
        }

        public static IPortion GetParagraphPortion(SCPresentation presentation, SlideElementQuery elementRequest)
        {
            IAutoShape autoShape = (IAutoShape)presentation.Slides[elementRequest.SlideIndex].Shapes.First(sp => sp.Id == elementRequest.ShapeId);
            
            return autoShape.TextBox.Paragraphs[elementRequest.ParagraphIndex].Portions[elementRequest.PortionIndex];
        }

        public static MemoryStream ToResizeableStream(this byte[] byteArray)
        {
            var stream = new MemoryStream();
            stream.Write(byteArray, 0, byteArray.Length);

            return stream;
        }
    }
}