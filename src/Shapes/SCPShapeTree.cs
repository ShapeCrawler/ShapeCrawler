using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

// ReSharper disable once InconsistentNaming
internal readonly ref struct SCPShapeTree(P.ShapeTree pShapeTree)
{
    internal void Add(OpenXmlElement openXmlElement)
    {
        var id = pShapeTree.Descendants<P.NonVisualDrawingProperties>().Select(s => s.Id!.Value).Max() + 1;
        var existingShapeNames = pShapeTree.Descendants<P.NonVisualDrawingProperties>().Select(s => s.Name!.Value!);
        var pShapeCopy = openXmlElement.CloneNode(true);
        pShapeCopy.NonVisualDrawingProperties().Id = new UInt32Value(id);
        pShapeTree.AppendChild(pShapeCopy);
        var copyName = pShapeCopy.NonVisualDrawingProperties().Name!.Value!;
        if (existingShapeNames.Any(existingShapeName => existingShapeName == copyName))
        {
            var currentShapeCollectionSuffixes = existingShapeNames
                .Where(c => c.StartsWith(copyName, StringComparison.InvariantCulture))
                .Select(c => c[copyName.Length..])
                .ToArray();

            // We will try to check numeric suffixes only.
            var numericSuffixes = new List<int>();

            foreach (var currentSuffix in currentShapeCollectionSuffixes)
            {
                if (int.TryParse(currentSuffix, out var numericSuffix))
                {
                    numericSuffixes.Add(numericSuffix);
                }
            }

            numericSuffixes.Sort();
            var lastSuffix = numericSuffixes.LastOrDefault() + 1;
            pShapeCopy.NonVisualDrawingProperties().Name = copyName + " " + lastSuffix;
        }
    }

    internal P.Shape? ReferencedPShapeOrNull(P.PlaceholderShape pPlaceholder)
    {
        var pShapes = pShapeTree.Elements<P.Shape>();
        foreach (var layoutPShape in pShapes)
        {
            var layoutPPlaceholder = layoutPShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>();
            if (layoutPPlaceholder == null)
            {
                continue;
            }

            if (pPlaceholder.Type?.Value == layoutPPlaceholder.Type?.Value &&
                pPlaceholder.Index?.Value == layoutPPlaceholder.Index?.Value)
            {
                return layoutPShape;
            }
        }

        if (pPlaceholder.Type?.Value is not null)
        {
            var byType = pShapes.FirstOrDefault(layoutPShape =>
                layoutPShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                    .GetFirstChild<P.PlaceholderShape>()?.Type?.Value == pPlaceholder.Type.Value);

            if (byType != null)
            {
                return byType;
            }
        }

        return null;
    }
}