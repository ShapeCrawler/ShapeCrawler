using System;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Assets;
using ShapeCrawler.Units;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

/// <summary>
///     Represents a connection shape.
/// </summary>
internal sealed class ConnectionShape(SlidePart slidePart, NewShapeProperties newShapeProperties)
{
    /// <summary>
    ///     Creates a connection shape with the specified coordinates.
    /// </summary>
    internal void Create(int startPointX, int startPointY, int endPointX, int endPointY)
    {
        var xml = new AssetCollection(Assembly.GetExecutingAssembly()).StringOf("new line.xml");
        var pConnectionShape = new P.ConnectionShape(xml);
        slidePart.Slide!.CommonSlideData!.ShapeTree!.Append(pConnectionShape);

        var x = Math.Min(startPointX, endPointX);
        var y = Math.Min(startPointY, endPointY);
        var cx = Math.Abs(endPointX - startPointX);
        var cy = Math.Abs(endPointY - startPointY);
        var flipH = startPointX > endPointX;
        var flipV = startPointY > endPointY;

        pConnectionShape.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!.Id = (uint)newShapeProperties.Id();

        var xEmu = new Points(x).AsEmus();
        var yEmu = new Points(y).AsEmus();
        var cxEmu = new Points(cx).AsEmus();
        var cyEmu = new Points(cy).AsEmus();
        var aXfrm = pConnectionShape.ShapeProperties!.Transform2D!;
        aXfrm.Offset!.X = xEmu;
        aXfrm.Offset!.Y = yEmu;
        aXfrm.Extents!.Cx = cxEmu;
        aXfrm.Extents!.Cy = cyEmu;
        aXfrm.HorizontalFlip = new BooleanValue(flipH);
        aXfrm.VerticalFlip = new BooleanValue(flipV);
    }
}