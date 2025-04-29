using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.DevTests.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlideParagraphAttribute(
    string caseName,
    string pptxName,
    int slideNumber,
    string shapeName,
    int paragraphNumber,
    object expectedResult)
    : Attribute, ITestBuilder
{
    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new Presentation(SCTest.TestAsset(pptxName));
        var paragraph = pres.Slides[slideNumber - 1].Shapes.Shape(shapeName).TextBox
            .Paragraphs[paragraphNumber - 1];

        var parameters = new TestCaseParameters(new[] { paragraph, expectedResult });

        if (!string.IsNullOrEmpty(caseName))
        {
            parameters.TestName = caseName;
        }

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}