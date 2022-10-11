using System.Collections.Generic;
using System.Reflection;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Helpers;

public class SlideDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly int slideNumber;
    private readonly object expectedResult;

    public SlideDataAttribute(string pptxFile, int slideNumber, object expectedResult)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.expectedResult = expectedResult;
    }
        
    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var pptxStream = TestHelper.GetStream(this.pptxFile);
        var pres = SCPresentation.Open(pptxStream);
        var slide = pres.Slides[this.slideNumber - 1];
     
        yield return new[] { slide, this.expectedResult };
    }
}