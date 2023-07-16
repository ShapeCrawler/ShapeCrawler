using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ShapeCrawler.Shapes;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Unit.Helpers.Attributes;

public class LayoutShapeDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly int slideNumber;
    private readonly string shapeName;

    public LayoutShapeDataAttribute(string pptxFile, int slideNumber, string shapeName)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
    }

    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var assembly = Assembly.GetExecutingAssembly();
        var stream = assembly.GetResourceStream(this.pptxFile);
        var pptxStream = new MemoryStream();
        stream.CopyTo(pptxStream);
        pptxStream.Position = 0;
        var pres = SCPresentation.Open(pptxStream);
        var layout = pres.SlideMasters[0].SlideLayouts[this.slideNumber - 1];
        var shape = layout.Shapes.GetByName<IShape>(this.shapeName);

        yield return new object[] { shape };
    }
}