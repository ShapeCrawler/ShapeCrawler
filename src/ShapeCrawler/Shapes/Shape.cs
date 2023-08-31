using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed record Shape
{
    private readonly OpenXmlCompositeElement sdkPShapeTreeElement;
    private readonly Position position;
    private readonly ShapeSize size;
    private const string customDataElementName = "ctd";

    internal Shape(OpenXmlCompositeElement sdkPShapeTreeElement)
    {
        this.sdkPShapeTreeElement = sdkPShapeTreeElement;
        this.position = new Position(this.sdkPShapeTreeElement);
        this.size = new ShapeSize(sdkPShapeTreeElement);
    }
    
    internal int X() => this.position.X();

    internal void UpdateX(int value) => this.position.UpdateX(value);

    internal int Y() => this.position.Y();

    internal void UpdateY(int value) => this.position.UpdateY(value);

    internal int Width() => this.size.Width();

    internal void UpdateWidth(int pixels) => this.size.UpdateWidth(pixels);

    internal int Height() => this.size.Height();

    internal void UpdateHeight(int pixels) => this.size.UpdateHeight(pixels);

    internal int Id() => (int)this.sdkPShapeTreeElement.GetNonVisualDrawingProperties().Id!.Value!;

    internal string Name() => this.sdkPShapeTreeElement.GetNonVisualDrawingProperties().Name!.Value!;

    internal bool Hidden()
    {
        var parsedHiddenValue = this.sdkPShapeTreeElement.GetNonVisualDrawingProperties().Hidden?.Value;
        
        return parsedHiddenValue is true;
    }

    internal SCGeometry GeometryType()
    {
        throw new System.NotImplementedException();
    }

    public string? CustomData()
    {
        const string pattern = @$"<{customDataElementName}>(.*)<\/{customDataElementName}>";

#if NETSTANDARD2_0
        var regex = new Regex(pattern, RegexOptions.None, TimeSpan.FromSeconds(100));
#else
        var regex = new Regex(pattern, RegexOptions.NonBacktracking);
#endif

        var elementText = regex.Match(this.sdkPShapeTreeElement.InnerXml).Groups[1];
        if (elementText.Value.Length == 0)
        {
            return null;
        }

        return elementText.Value;
    }

    public void UpdateCustomData(string? value)
    {
        var customDataElement =
            $@"<{customDataElementName}>{value}</{customDataElementName}>";
        this.sdkPShapeTreeElement.InnerXml += customDataElement;
    }
}