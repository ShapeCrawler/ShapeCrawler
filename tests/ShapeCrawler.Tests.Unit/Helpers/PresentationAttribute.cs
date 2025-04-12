using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;
using ShapeCrawler.Presentations;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class PresentationAttribute(string pptxName) : Attribute, ITestBuilder
{
    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        IPresentation pres;
        if (pptxName == "new")
        {
            pres = new Presentation();
        }
        else
        {
            var pptxStream = SCTest.TestAsset(pptxName);
            pres = new Presentation(pptxStream);
        }

        var parameters = new TestCaseParameters(new object[] { pres });

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}