using System.Collections.Generic;
using System.Reflection;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using ShapeCrawler.Shapes;
using ShapeCrawler.Tests.Shared;
using Xunit.Sdk;

namespace ShapeCrawler.Tests.Unit.Helpers.Attributes;

public class SlideShapeDataAttribute : DataAttribute
{
    private readonly string pptxFile;
    private readonly int slideNumber;
    private readonly string shapeName;
    private readonly int shapeId;
    private readonly object expectedResult;
    private readonly string displayName;

    public SlideShapeDataAttribute(string pptxFile, int slideNumber, string shapeName, object expectedResult)
    : this(pptxFile, slideNumber, shapeName)
    {
        this.expectedResult = expectedResult;
    }

    public SlideShapeDataAttribute(string displayName, string pptxFile, int slideNumber, string shapeName)
        : this(pptxFile, slideNumber, shapeName)
    {
        this.displayName = displayName;
    }
    
    public SlideShapeDataAttribute(string pptxFile, int slideNumber, string shapeName)
    {
        this.pptxFile = pptxFile;
        this.slideNumber = slideNumber;
        this.shapeName = shapeName;
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
        var helperAssets = new List<string> { "autoshape-grouping.pptx", "001.pptx", "table-case001.pptx", "autoshape-case005_text-frame.pptx" };
        var pptxStream = helperAssets.Contains(this.pptxFile) ? Tests.Shared.TestHelper.GetStream(this.pptxFile) : TestHelperOld.GetStream(this.pptxFile);
        var pres = SCPresentation.Open(pptxStream);
        var slide = pres.Slides[this.slideNumber - 1];
        var shape = this.shapeName != null
            ? slide.Shapes.GetByName<IShape>(this.shapeName)
            : slide.Shapes.GetById<IShape>(this.shapeId);

        var input = new List<object>();

        if (this.displayName != null)
        {
            input.Add(this.displayName);
        }
        
        input.Add(shape);
        
        if (this.expectedResult != null)
        {
            input.Add(this.expectedResult);
        }

        yield return input.ToArray();
    }
}