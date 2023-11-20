using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.Tests.Unit.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlideQueryPortionAttribute : Attribute, ITestBuilder
{
    private readonly string pptxName;
    private readonly int slideNumber;
    private readonly string shapeName;
    private readonly int shapeId;
    private readonly int paragraphNumber;
    private readonly int portionNumber;
    private readonly object expectedResult;

    public SlideQueryPortionAttribute(string pptxName, int slideNumber, string shapeName, int paragraphNumber, int portionNumber)
    {
        this.pptxName = pptxName;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
        this.paragraphNumber = paragraphNumber;
        this.portionNumber = portionNumber;
    }
    
    public SlideQueryPortionAttribute(string pptxName, int slideNumber, int shapeId, int paragraphNumber, int portionNumber)
    {
        this.pptxName = pptxName;
        this.slideNumber = slideNumber;
        this.shapeId = shapeId;
        this.paragraphNumber = paragraphNumber;
        this.portionNumber = portionNumber;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pres = new Presentation(SCTest.StreamOf(this.pptxName));
        var portionQuery = this.shapeName == null 
            ? new TestSlidePortionQuery(this.slideNumber, this.shapeId, this.paragraphNumber, this.portionNumber) 
            : new TestSlidePortionQuery(this.slideNumber, this.shapeName, this.paragraphNumber, this.portionNumber);

        var parameters = new TestCaseParameters(new object[] { pres, portionQuery });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}