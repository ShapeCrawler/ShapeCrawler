using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class LayoutShapeAttribute : Attribute, ITestBuilder
{
    private readonly string pptxName;
    private readonly int slideLayoutNumber;
    private readonly string shapeName;
    private readonly object? expectedResult;

    public LayoutShapeAttribute(string pptxName, int slideLayoutNumber, string shapeName)
    {
        this.pptxName = pptxName;
        this.slideLayoutNumber = slideLayoutNumber;
        this.shapeName = shapeName;
    }
    
    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pptxStream = SCTest.StreamOf(this.pptxName);
        var pres = new Presentation(pptxStream);
        var shape = pres.SlideMasters[0].SlideLayouts[this.slideLayoutNumber - 1].Shapes.GetByName(this.shapeName);

        var parameters = new TestCaseParameters(new[] { shape });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}