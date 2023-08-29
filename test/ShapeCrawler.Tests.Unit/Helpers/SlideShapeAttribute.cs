using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlideShapeAttribute : Attribute, ITestBuilder
{
    private readonly int slideNumber;
    private readonly int shapeId;
    private readonly object expectedResult;
    private readonly string pptxName;

    public SlideShapeAttribute(string pptxName, int slideNumber, int shapeId, object expectedResult)
    {
        this.pptxName = pptxName;
        this.slideNumber = slideNumber;
        this.shapeId = shapeId;
        this.expectedResult = expectedResult;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pptxStream = SCTest.StreamOf(this.pptxName);
        var pres = new SCPresentation(pptxStream);
        var shape = pres.Slides[this.slideNumber - 1].Shapes.GetById<IShape>(this.shapeId);

        var parameters = new TestCaseParameters(new[] { shape, this.expectedResult });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}