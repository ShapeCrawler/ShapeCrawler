using System.Collections.Generic;
using System.Reflection;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Unit.Helpers.Attributes;

public class SlideParagraphDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly int slideNumber;
    private readonly string shapeName;
    private readonly object expectedResult;
    private readonly int paraNumber;

    public SlideParagraphDataAttribute(
        string pptxFile, 
        int slideNumber, 
        string shapeName, 
        int paraNumber,
        object expectedResult)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
        this.paraNumber = paraNumber;
        this.expectedResult = expectedResult;
    }

    public override IEnumerable<object[]> GetData(MethodInfo testMethod)
    {
        var pptxStream = TestHelper.GetStream(this.pptxFile);
        var pres = new Presentation(pptxStream);
        var slide = pres.Slides[this.slideNumber - 1];
        var shape = slide.Shapes.GetByName(this.shapeName);
        var paragraph = shape.TextFrame!.Paragraphs[this.paraNumber - 1];

        yield return new[] { paragraph, this.expectedResult };
    }
}