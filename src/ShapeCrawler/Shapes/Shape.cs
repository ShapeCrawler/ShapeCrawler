using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed class Shape
{
    private readonly OpenXmlElement pShapeTreeElement;
    private readonly Position position;
    private readonly ShapeSize size;
    private const string customDataElementName = "ctd";

    internal Shape(OpenXmlElement pShapeTreeElement)
    {
        this.pShapeTreeElement = pShapeTreeElement;
        this.position = new Position(this.pShapeTreeElement);
        this.size = new ShapeSize(pShapeTreeElement);
    }
    
    internal int X() => this.position.X();

    internal void UpdateX(int value) => this.position.UpdateX(value);

    internal int Y() => this.position.Y();

    internal void UpdateY(int value) => this.position.UpdateY(value);

    internal int Width() => this.size.Width();

    internal void UpdateWidth(int pixels) => this.size.UpdateWidth(pixels);

    internal int Height() => this.size.Height();

    internal void UpdateHeight(int pixels) => this.size.UpdateHeight(pixels);

    internal int Id() => (int)this.pShapeTreeElement.GetNonVisualDrawingProperties().Id!.Value!;

    internal string Name() => this.pShapeTreeElement.GetNonVisualDrawingProperties().Name!.Value!;

    internal bool Hidden()
    {
        var parsedHiddenValue = this.pShapeTreeElement.GetNonVisualDrawingProperties().Hidden?.Value;
        
        return parsedHiddenValue is true;
    }

    internal SCGeometry GeometryType() => SCGeometry.Rectangle;

    internal string? CustomData()
    {
        const string pattern = @$"<{customDataElementName}>(.*)<\/{customDataElementName}>";

#if NETSTANDARD2_0
        var regex = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(100));
#else
        var regex = new Regex(pattern, RegexOptions.NonBacktracking);
#endif

        var elementText = regex.Match(this.pShapeTreeElement.InnerXml).Groups[1];
        if (elementText.Value.Length == 0)
        {
            return null;
        }

        return elementText.Value;
    }

    internal void UpdateCustomData(string? value)
    {
        var customDataElement =
            $@"<{customDataElementName}>{value}</{customDataElementName}>";
        this.pShapeTreeElement.InnerXml += customDataElement;
    }
}