using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.DevTests.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlidePortionAttribute(
    string pptxName,
    int slide,
    int shapeId,
    int paragraph,
    int portion,
    object expectedResult)
    : Attribute, ITestBuilder
{
    private readonly string testName;

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

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new Presentation(SCTest.TestAsset(pptxName));
        var portion1 = pres.Slides[slide - 1].Shapes.GetById<IShape>(shapeId).TextBox
            .Paragraphs[paragraph - 1].Portions[portion - 1];

        var parameters = new TestCaseParameters(new[] { portion1, expectedResult });

        if (!string.IsNullOrEmpty(this.testName))
        {
            parameters.TestName = this.testName;
        }

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}