using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Drawing;

internal sealed class CellFill : SCShapeFill
{
    internal CellFill(
        ISlideStructure slideStructure, 
        TypedOpenXmlCompositeElement cellProperties, 
        TypedOpenXmlPart slideTypedOpenXmlPart,
        List<ImagePart> imageParts)
        : base(slideStructure, cellProperties, slideTypedOpenXmlPart, imageParts)
    {
    }
}