namespace ShapeCrawler.Tests.Unit.Helpers;

public class TestMasterPortionQuery : TestPortionQuery
{
    private readonly string shapeName;
    private readonly int paragraphNumber;
    private readonly int portionNumber;

    public TestMasterPortionQuery(string shapeName, int paragraphNumber, int portionNumber)
    {
        this.shapeName = shapeName;
        this.paragraphNumber = paragraphNumber;
        this.portionNumber = portionNumber;
    }

    public override IParagraphPortion Get(IPresentation pres)
    {
        return pres.SlideMasters[0].Shapes.GetByName<IShape>(this.shapeName).TextFrame!
            .Paragraphs[this.paragraphNumber - 1].Portions[this.portionNumber - 1];
    }
}