using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;
using ShapeCrawler.Presentations;

namespace ShapeCrawler.DevTests.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class MasterPortionAttribute(string pptxName, string shapeName, int paragraphNumber, int portionNumber)
    : Attribute, ITestBuilder
{
    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new Presentation(SCTest.TestAsset(pptxName));
        var portionQuery = new TestMasterPortionQuery(shapeName, paragraphNumber, portionNumber);

        var parameters = new TestCaseParameters(new object[] { pres, portionQuery });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}