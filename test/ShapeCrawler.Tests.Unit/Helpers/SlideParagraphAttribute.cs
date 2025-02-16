using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlideParagraphAttribute : Attribute, ITestBuilder
{
    private readonly string caseName;
    private readonly string pptxName;
    private readonly int slideNumber;
    private readonly string shapeName;
    private readonly int paragraphNumber;
    private readonly object expectedResult;

    public SlideParagraphAttribute(
        string caseName,
        string pptxName,
        int slideNumber,
        string shapeName,
        int paragraphNumber,
        object expectedResult)
    {
        this.caseName = caseName;
        this.pptxName = pptxName;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
        this.paragraphNumber = paragraphNumber;
        this.expectedResult = expectedResult;
    }
    
    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new Presentation(SCTest.TestAsset(this.pptxName));
        var paragraph = pres.Slides[this.slideNumber - 1].Shapes.GetByName(this.shapeName).TextBox
            .Paragraphs[this.paragraphNumber - 1];

        var parameters = new TestCaseParameters(new[] { paragraph, this.expectedResult });

        if (!string.IsNullOrEmpty(this.caseName))
        {
            parameters.TestName = this.caseName;
        }

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}