using System.Linq;

namespace ShapeCrawler.Tests.Unit.Helpers;

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
    public int? SlideNumber { get; set; }


    public IShape GetAutoShape()
    {
        var slideIndex = 0;
        if (this.SlideNumber != null)
        {
            slideIndex = this.SlideNumber.Value - 1;
        }
        else
        {
            slideIndex = this.SlideIndex;
        }
            
        var shapes = this.Presentation.Slides[slideIndex].Shapes;
        return this.ShapeName != null ? shapes.GetByName<IShape>(this.ShapeName) : shapes.GetById<IShape>(this.ShapeId!.Value);
    }
        
    public IParagraph GetParagraph()
    {
        var paragraphIndex = this.ParagraphIndex ?? this.ParagraphNumber - 1;
        var autoShape = Presentation.Slides[SlideIndex]
            .Shapes.First(sp => sp.Id == ShapeId) as IShape;
        return autoShape.TextFrame.Paragraphs[paragraphIndex];
    }

    public IParagraphPortion GetParagraphPortion()
    {
        var shapes = this.Presentation.Slides[this.SlideIndex].Shapes;
        var autoShape = this.ShapeId != null 
            ? shapes.GetById<IShape>(this.ShapeId.Value) 
            : shapes.GetByName<IShape>(this.ShapeName);

        var paragraphIndex = this.ParagraphIndex ?? this.ParagraphNumber - 1;
        var portionIndex = this.PortionIndex ?? this.PortionNumber - 1;
            
        return autoShape.TextFrame!.Paragraphs[paragraphIndex].Portions[portionIndex];
    }

    public IFontColor GetTestColorFormat()
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
            ? shapes.GetById<IShape>(this.ShapeId.Value) 
            : shapes.GetByName<IShape>(this.ShapeName);
            
        var paragraphIndex = this.ParagraphIndex ?? this.ParagraphNumber - 1;
        var portionIndex = this.PortionIndex ?? this.PortionNumber - 1;

        return autoShape.TextFrame!.Paragraphs[paragraphIndex].Portions[portionIndex].Font.Color;
    }
}