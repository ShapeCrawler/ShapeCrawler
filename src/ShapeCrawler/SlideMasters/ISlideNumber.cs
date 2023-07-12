using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler;

public interface ISlideNumber
{
    ISlideNumberFont Font { get; }
}

internal class SCSlideNumber : ISlideNumber
{
    private readonly P.ShapeTree pShapeTree;

    public SCSlideNumber(P.ShapeTree pShapeTree)
    {
        this.pShapeTree = pShapeTree;
        var pSldNum = this.pShapeTree.Elements<P.Shape>().FirstOrDefault(shape => shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.PlaceholderShape!.Type!.Value == P.PlaceholderValues.SlideNumber);
        var aDefaultRunProperties = pSldNum!.TextBody!.ListStyle!.Level1ParagraphProperties?.GetFirstChild<A.DefaultRunProperties>()!; 
        this.Font = new SCSlideNumberFont(aDefaultRunProperties);
    }

    public ISlideNumberFont Font { get; }
}