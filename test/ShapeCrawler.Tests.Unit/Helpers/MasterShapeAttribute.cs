using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class MasterShapeAttribute : Attribute, ITestBuilder
{
    private readonly string pptxName;
    private readonly string shapeName;
    private readonly object? expectedResult;

    public MasterShapeAttribute(string pptxName, string shapeName)
    {
        this.pptxName = pptxName;
        this.shapeName = shapeName;
    }
    
    public MasterShapeAttribute(string pptxName, string shapeName, object expectedResult)
    {
        this.pptxName = pptxName;
        this.shapeName = shapeName;
        this.expectedResult = expectedResult;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pptxStream = SCTest.TestAsset(this.pptxName);
        var pres = new Presentation(pptxStream);
        var shape = pres.SlideMasters[0].Shapes.GetByName(this.shapeName);

        var parameters = this.expectedResult != null
            ? new TestCaseParameters(new[] { shape, this.expectedResult })
            : new TestCaseParameters(new[] { shape });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}