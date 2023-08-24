﻿using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Shapes;

internal sealed record Shape
{
    private readonly OpenXmlCompositeElement pShapeTreeChild;
    private readonly Position position;
    private const string customDataElementName = "ctd";

    internal Shape(OpenXmlCompositeElement pShapeTreeChild)
    {
        this.pShapeTreeChild = pShapeTreeChild;
        
        var aOffset = pShapeTreeChild.Descendants<A.Offset>().First();
        this.position = new Position(aOffset);
    }
    
    internal int X() => this.position.X();

    internal void UpdateX(int value) => this.position.UpdateX(value);

    internal int Y() => this.position.Y();

    internal void UpdateY(int value) => this.position.UpdateY(value);

    internal int Width()
    {
        throw new System.NotImplementedException();
    }

    internal void UpdateWidth(int value)
    {
        throw new System.NotImplementedException();
    }

    internal int Height()
    {
        throw new System.NotImplementedException();
    }

    internal void UpdateHeight(int value)
    {
        throw new System.NotImplementedException();
    }

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

        var elementText = regex.Match(this.pShapeTreeChild.InnerXml).Groups[1];
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
        this.pShapeTreeChild.InnerXml += customDataElement;
    }
}