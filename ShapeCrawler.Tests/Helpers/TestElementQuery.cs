using System.Linq;
using ShapeCrawler.Drawing;

namespace ShapeCrawler.Tests.Helpers
{
    public enum Location
    {
        Slide = 1,
        SlideLayout = 2,
        SlideMaster = 3
    }
    
    public class TestElementQuery
    {
        public int SlideIndex { get; set; }
        public int? ShapeId { get; set; }
        public int? ParagraphIndex { get; set; }
        public int ParagraphNumber { get; set; }
        public int? PortionIndex { get; set; }
        public int PortionNumber { get; set; }
        public IPresentation Presentation { get; set; }
        public string ShapeName { get; set; }
        public Location Location { get; set; }
        public int SlideMasterNumber { get; set; }
        public int SlideLayoutNumber { get; set; }
        
        public IParagraph GetParagraph()
        {
            var paragraphIndex = this.ParagraphIndex ?? this.ParagraphNumber - 1;
            var autoShape = Presentation.Slides[SlideIndex]
                .Shapes.First(sp => sp.Id == ShapeId) as IAutoShape;
            return autoShape.TextBox.Paragraphs[paragraphIndex];
        }

        public IPortion GetParagraphPortion()
        {
            var shapes = this.Presentation.Slides[this.SlideIndex].Shapes;
            var autoShape = this.ShapeId != null 
                ? shapes.GetById<IAutoShape>(this.ShapeId.Value) 
                : shapes.GetByName<IAutoShape>(this.ShapeName);

            var paragraphIndex = this.ParagraphIndex ?? this.ParagraphNumber - 1;
            var portionIndex = this.PortionIndex ?? this.PortionNumber - 1;
            
            return autoShape.TextBox!.Paragraphs[paragraphIndex].Portions[portionIndex];
        }

        public IColorFormat GetTestColorFormat()
        {
            var shapes = this.Location switch
            {
                Location.Slide => this.Presentation.Slides[this.SlideIndex].Shapes,
                Location.SlideLayout => this.Presentation.SlideMasters[this.SlideMasterNumber - 1]
                    .SlideLayouts[this.SlideLayoutNumber - 1]
                    .Shapes,
                _ => this.Presentation.SlideMasters[this.SlideMasterNumber - 1].Shapes
            };

            var autoShape = this.ShapeId != null 
                ? shapes.GetById<IAutoShape>(this.ShapeId.Value) 
                : shapes.GetByName<IAutoShape>(this.ShapeName);
            
            var paragraphIndex = this.ParagraphIndex ?? this.ParagraphNumber - 1;
            var portionIndex = this.PortionIndex ?? this.PortionNumber - 1;

            return autoShape.TextBox!.Paragraphs[paragraphIndex].Portions[portionIndex].Font.ColorFormat;
        }
    }
}