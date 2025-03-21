using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;
using ShapeCrawler.Presentations;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class LayoutShapeAttribute(string pptxName, int slideLayoutNumber, string shapeName) : Attribute, ITestBuilder
{
    private readonly object? expectedResult;

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pptxStream = SCTest.TestAsset(pptxName);
        var pres = new Presentation(pptxStream);
        var shape = pres.SlideMasters[0].SlideLayouts[slideLayoutNumber - 1].Shapes.GetByName(shapeName);

        var parameters = new TestCaseParameters(new[] { shape });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}