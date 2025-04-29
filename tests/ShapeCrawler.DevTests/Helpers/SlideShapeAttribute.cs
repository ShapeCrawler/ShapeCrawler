using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Builders;

namespace ShapeCrawler.DevTests.Helpers;

[AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
public class SlideShapeAttribute : Attribute, ITestBuilder
{
    private readonly int slideNumber;
    private readonly int? shapeId;
    private readonly object? expectedResult;
    private readonly string pptxName;
    private readonly string shapeName;

    public SlideShapeAttribute(string pptxName, int slideNumber, string shapeName)
    {
        this.pptxName = pptxName;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
    }
    
    public SlideShapeAttribute(string pptxName, int slideNumber, int shapeId, object expectedResult)
    {
        this.pptxName = pptxName;
        this.slideNumber = slideNumber;
        this.shapeId = shapeId;
        this.expectedResult = expectedResult;
    }
    
    public SlideShapeAttribute(string pptxName, int slideNumber, string shapeName, object expectedResult)
    {
        this.pptxName = pptxName;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
        this.expectedResult = expectedResult;
    }

    public IEnumerable<TestMethod> BuildFrom(IMethodInfo method, Test suite)
    {
        var pptxStream = SCTest.TestAsset(this.pptxName);
        var pres = new Presentation(pptxStream);
        var shape = this.shapeId.HasValue 
            ? pres.Slides[this.slideNumber - 1].Shapes.GetById<IShape>(this.shapeId.Value) 
            : pres.Slides[this.slideNumber - 1].Shapes.Shape<IShape>(this.shapeName);

        var parameters = this.expectedResult != null
            ? new TestCaseParameters(new[] { shape, this.expectedResult })
            : new TestCaseParameters(new[] { shape });
        
        yield return new NUnitTestCaseBuilder().BuildTestMethod(method, suite, parameters);
    }
}