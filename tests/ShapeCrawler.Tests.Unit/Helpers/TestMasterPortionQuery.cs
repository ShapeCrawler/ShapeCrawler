namespace ShapeCrawler.Tests.Unit.Helpers;

public class TestMasterPortionQuery(string shapeName, int paragraphNumber, int portionNumber) : TestPortionQuery
{
    public override IParagraphPortion Get(IPresentation pres)
    {
        return pres.SlideMasters[0].Shapes.GetByName<IShape>(shapeName).TextBox!
            .Paragraphs[paragraphNumber - 1].Portions[portionNumber - 1];
    }
}