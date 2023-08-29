using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class MasterPortionAttribute : Attribute, ITestBuilder
{
    private readonly string pptxName;
    private readonly string shapeName;
    private readonly int paragraphNumber;
    private readonly int portionNumber;

    public MasterPortionAttribute(string pptxName, string shapeName, int paragraphNumber, int portionNumber)
    {
        this.pptxName = pptxName;
        this.shapeName = shapeName;
        this.paragraphNumber = paragraphNumber;
        this.portionNumber = portionNumber;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new SCPresentation(SCTest.StreamOf(this.pptxName));
        var portionQuery = new TestMasterPortionQuery(this.shapeName, this.paragraphNumber, this.portionNumber);

        var parameters = new TestCaseParameters(new object[] { pres, portionQuery });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}