using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class PresentationAttribute : Attribute, ITestBuilder
{
    private readonly string pptxName;

    public PresentationAttribute(string pptxName)
    {
        this.pptxName = pptxName;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        IPresentation pres;
        if (this.pptxName == "new")
        {
            pres = new SCPresentation();
        }
        else
        {
            var pptxStream = SCTest.StreamOf(this.pptxName);
            pres = new SCPresentation(pptxStream);
        }

        var parameters = new TestCaseParameters(new object[] { pres });

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}