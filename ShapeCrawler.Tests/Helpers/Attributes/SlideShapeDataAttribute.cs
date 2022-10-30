using System.Collections.Generic;
using System.Reflection;
using ShapeCrawler.Shapes;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Helpers.Attributes;

public class SlideShapeDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly int slideNumber;
    private readonly string shapeName;
    private readonly int shapeId;
    private readonly object expectedResult;

    public SlideShapeDataAttribute(string pptxFile, int slideNumber, string shapeName)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
    }

    public SlideShapeDataAttribute(string pptxFile, int slideNumber, int shapeId)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.shapeId = shapeId;
    }
    
    public SlideShapeDataAttribute(string pptxFile, int slideNumber, int shapeId, object expectedResult)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.shapeId = shapeId;
        this.expectedResult = expectedResult;
    }

    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var pptxStream = TestHelper.GetStream(this.pptxFile);
        var pres = SCPresentation.Open(pptxStream);
        var slide = pres.Slides[this.slideNumber - 1];
        var shape = this.shapeName != null
            ? slide.Shapes.GetByName<IShape>(this.shapeName)
            : slide.Shapes.GetById<IShape>(this.shapeId);

        if (this.expectedResult != null)
        {
            yield return new object[] { shape, this.expectedResult };
        }
        else
        {
            yield return new object[] { shape };
        }
    }
}