namespace ShapeCrawler.Tests.Unit.Helpers;

public class TestSlidePortionQuery : TestPortionQuery
{
    private readonly int slideNumber;
    private readonly string shapeName;
    private readonly int shapeId;
    private readonly int paragraphNumber;
    private readonly int portionNumber;

    public TestSlidePortionQuery(int slideNumber, string shapeName, int paragraphNumber, int portionNumber)
    {
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
        this.paragraphNumber = paragraphNumber;
        this.portionNumber = portionNumber;
    }

    public TestSlidePortionQuery(int slideNumber, int shapeId, int paragraphNumber, int portionNumber)
    {
        this.slideNumber = slideNumber;
        this.shapeId = shapeId;
        this.paragraphNumber = paragraphNumber;
        this.portionNumber = portionNumber;
    }

    public override IParagraphPortion Get(IPresentation pres)
    {
        var shapes = pres.Slides[this.slideNumber - 1].Shapes;
        var shape = this.shapeName == null
            ? shapes.GetById<IShape>(this.shapeId)
            : shapes.GetByName<IShape>(this.shapeName);

        return shape.TextBox!.Paragraphs[this.paragraphNumber - 1].Portions[this.portionNumber - 1];
    }
}