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

    public SlideShapeDataAttribute(string pptxFile, int slideNumber, string shapeName)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
    }

    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var pptxStream = TestHelper.GetStream(this.pptxFile);
        var pres = SCPresentation.Open(pptxStream);
        var slide = pres.Slides[this.slideNumber - 1];
        var shape = slide.Shapes.GetByName<IShape>(this.shapeName);

        yield return new object[] { shape };
    }
}