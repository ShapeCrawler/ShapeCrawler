namespace ShapeCrawler.Tests.Unit.Helpers;

public class TestSlidePortionQuery(int slideNumber, string shapeName, int paragraphNumber, int portionNumber)
    : TestPortionQuery
{
    private readonly int shapeId;

    public TestSlidePortionQuery(int slideNumber, int shapeId, int paragraphNumber, int portionNumber) : this(slideNumber, null, paragraphNumber, portionNumber)
    {
        this.shapeId = shapeId;
    }

    public override IParagraphPortion Get(IPresentation pres)
    {
        var shapes = pres.Slides[slideNumber - 1].Shapes;
        var shape = shapeName == null
            ? shapes.GetById<IShape>(this.shapeId)
            : shapes.GetByName<IShape>(shapeName);

        return shape.TextBox!.Paragraphs[paragraphNumber - 1].Portions[portionNumber - 1];
    }
}