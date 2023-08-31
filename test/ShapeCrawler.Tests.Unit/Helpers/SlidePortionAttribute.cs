using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlidePortionAttribute : Attribute, ITestBuilder
{
    private readonly string pptxName;
    private readonly int slide;
    private readonly int shapeId;
    private readonly int paragraph;
    private readonly int portion;
    private readonly string testName;
    private readonly object expectedResult;

    public SlidePortionAttribute(
        string testName,
        string pptxName,
        int slide,
        int shapeId,
        int paragraph,
        int portion,
        object expectedResult)
        : this(pptxName, slide, shapeId, paragraph, portion, expectedResult)
    {
        this.testName = testName;
    }

    public SlidePortionAttribute(
        string pptxName,
        int slide,
        int shapeId,
        int paragraph,
        int portion,
        object expectedResult)
    {
        this.pptxName = pptxName;
        this.slide = slide;
        this.shapeId = shapeId;
        this.paragraph = paragraph;
        this.portion = portion;
        this.expectedResult = expectedResult;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new SCPresentation(SCTest.StreamOf(this.pptxName));
        var portion = pres.Slides[this.slide - 1].Shapes.GetById<IShape>(this.shapeId).TextFrame
            .Paragraphs[this.paragraph - 1].Portions[this.portion - 1];

        var parameters = new TestCaseParameters(new[] { portion, this.expectedResult });

        if (!string.IsNullOrEmpty(this.testName))
        {
            parameters.TestName = this.testName;
        }

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}