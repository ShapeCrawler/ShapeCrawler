using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace ShapeCrawler.Drawing;

internal interface ICopyableShape
{
    void CopyTo(
        int id, 
        DocumentFormat.OpenXml.Presentation.ShapeTree pShapeTree, 
        IEnumerable<string> existingShapeNames,
        SlidePart targetSdkSlidePart);
}