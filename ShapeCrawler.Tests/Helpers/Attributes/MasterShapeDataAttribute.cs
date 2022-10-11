using System.Collections.Generic;
using System.Reflection;
using ShapeCrawler.Shapes;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Helpers.Attributes;

public class MasterShapeDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly string shapeName;

    public MasterShapeDataAttribute(string pptxFile, string shapeName)
    {
        this.pptxFile = pptxFile;
        this.shapeName = shapeName;
    }

    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var pptxStream = TestHelper.GetStream(this.pptxFile);
        var pres = SCPresentation.Open(pptxStream);
        var slideMaster = pres.SlideMasters[0];
        var shape = slideMaster.Shapes.GetByName<IShape>(this.shapeName);

        yield return new object[] { shape };
    }
}