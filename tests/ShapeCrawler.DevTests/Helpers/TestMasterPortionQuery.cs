namespace ShapeCrawler.DevTests.Helpers;

public class TestMasterPortionQuery(string shapeName, int paragraphNumber, int portionNumber) : TestPortionQuery
{
    public override IParagraphPortion Get(IPresentation pres)
    {
        return pres.MasterSlides[0].Shapes.Shape<IShape>(shapeName).TextBox!
            .Paragraphs[paragraphNumber - 1].Portions[portionNumber - 1];
    }
}