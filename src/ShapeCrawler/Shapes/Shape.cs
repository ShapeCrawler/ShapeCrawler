using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed record Shape
{
    private readonly OpenXmlCompositeElement pShapeTreeElement;
    private readonly Position position;
    private readonly ShapeSize size;
    private const string customDataElementName = "ctd";

    internal Shape(OpenXmlCompositeElement pShapeTreeElement)
    {
        this.pShapeTreeElement = pShapeTreeElement;
        
        var aOffset = pShapeTreeElement.Descendants<A.Offset>().First();
        this.position = new Position(aOffset);
        this.size = new ShapeSize(pShapeTreeElement.Descendants<A.Extents>().First());
    }
    
    internal int X() => this.position.X();

    internal void UpdateX(int value) => this.position.UpdateX(value);

    internal int Y() => this.position.Y();

    internal void UpdateY(int value) => this.position.UpdateY(value);

    internal int Width() => this.size.Width();

    internal void UpdateWidth(int pixels) => this.size.UpdateWidth(pixels);

    internal int Height() => this.size.Height();

    internal void UpdateHeight(int pixels) => this.size.UpdateHeight(pixels);

    internal int Id()
    {
        throw new System.NotImplementedException();
    }

    internal string Name()
    {
        throw new System.NotImplementedException();
    }

    internal bool Hidden()
    {
        throw new System.NotImplementedException();
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

        var elementText = regex.Match(this.pShapeTreeElement.InnerXml).Groups[1];
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
        this.pShapeTreeElement.InnerXml += customDataElement;
    }
}