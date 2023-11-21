using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ShapeCrawler.Tests.Shared;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Unit.Helpers.Attributes;

public class SlideDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly int slideNumber;
    private readonly object expectedResult;
    private readonly string testCaseLabel;

    public SlideDataAttribute(string testCaseLabel, string pptxFile, int slideNumber, object expectedResult)
    {
        this.testCaseLabel = testCaseLabel;
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.expectedResult = expectedResult;
    }
        
    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var pptx = SCTest.StreamOf(this.pptxFile);
        var pres = new Presentation(pptx);
        var slide = pres.Slides[this.slideNumber - 1];

        if (this.testCaseLabel == null)
        {
            yield return new[] { slide, this.expectedResult };    
        }
        else
        {
            yield return new[] { this.testCaseLabel, slide, this.expectedResult };
        }
    }
}