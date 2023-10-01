using System.Collections.Generic;
using System.Reflection;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using ShapeCrawler.Tests.Unit.Helpers;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Unit.Helpers.Attributes;

public class MasterShapeDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly string shapeName;
    private readonly object expectedResult;

    public MasterShapeDataAttribute(string pptxFile, string shapeName)
    {
        this.pptxFile = pptxFile;
        this.shapeName = shapeName;
    }
    
    public MasterShapeDataAttribute(string pptxFile, string shapeName, object expectedResult)
    {
        this.pptxFile = pptxFile;
        this.shapeName = shapeName;
        this.expectedResult = expectedResult;
    }

    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var pptxStream = SCTest.StreamOf(this.pptxFile);
        var pres = new Presentation(pptxStream);
        var slideMaster = pres.SlideMasters[0];
        var shape = slideMaster.Shapes.GetByName<IShape>(this.shapeName);

        if (this.expectedResult == null)
        {
            yield return new object[] { shape };    
        }
        else
        {
            yield return new object[] { shape, this.expectedResult };
        }
    }
}