using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;
using ShapeCrawler.Presentations;

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
            pres = new Presentation();
        }
        else
        {
            var pptxStream = SCTest.TestAsset(this.pptxName);
            pres = new Presentation(pptxStream);
        }

        var parameters = new TestCaseParameters(new object[] { pres });

        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}