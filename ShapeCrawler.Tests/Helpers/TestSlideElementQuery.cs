using System.Linq;

namespace ShapeCrawler.Tests.Helpers
{
    public class TestSlideElementQuery
    {
        public int SlideIndex { get; set; }
        public int? ShapeId { get; set; }
        public int ParagraphIndex { get; set; }
        public int PortionIndex { get; set; }
        public IPresentation Presentation { get; set; }
        public string ShapeName { get; set; }

        public IParagraph GetParagraph()
        {
            var autoShape = Presentation.Slides[SlideIndex]
                .Shapes.First(sp => sp.Id == ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[ParagraphIndex];
        }

        public IPortion GetParagraphPortion()
        {
            var shapes = this.Presentation.Slides[this.SlideIndex].Shapes;
            var autoShape = this.ShapeId != null 
                ? shapes.GetById<IAutoShape>(this.ShapeId.Value) 
                : shapes.GetByName<IAutoShape>(this.ShapeName);

            return autoShape.TextBox!.Paragraphs[this.ParagraphIndex].Portions[this.PortionIndex];
        }
    }
}