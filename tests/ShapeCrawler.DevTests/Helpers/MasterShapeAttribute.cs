using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.DevTests.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class MasterShapeAttribute(string pptxName, string shapeName, object expectedResult) : Attribute, ITestBuilder
{
    public MasterShapeAttribute(string pptxName, string shapeName) : this(pptxName, shapeName, null)
    {
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pptxStream = SCTest.TestAsset(pptxName);
        var pres = new Presentation(pptxStream);
        var shape = pres.SlideMasters[0].Shapes.Shape(shapeName);

        var parameters = expectedResult != null
            ? new TestCaseParameters(new[] { shape, expectedResult })
            : new TestCaseParameters(new[] { shape });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}