using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

// ReSharper disable PossibleMultipleEnumeration
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

    internal P.Shape? ReferencedPShapeOrNull(P.PlaceholderShape pPlaceholderShape)
    {
        var pShapes = GetShapesWithPlaceholder(pShapeTree);

        // First try to find shape by matching both type and index
        var matchByTypeAndIndex = FindShapeByTypeAndIndex(pShapes, pPlaceholderShape);
        if (matchByTypeAndIndex != null)
        {
            return matchByTypeAndIndex;
        }

        // If not found, try to find shape by type only
        return FindShapeByTypeOnly(pShapes, pPlaceholderShape);
    }

    private static IEnumerable<P.Shape> GetShapesWithPlaceholder(P.ShapeTree shapeTree)
    {
        return shapeTree.Elements<P.Shape>().Where(x =>
            x.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.GetFirstChild<P.PlaceholderShape>() != null);
    }

    private static P.Shape? FindShapeByTypeAndIndex(IEnumerable<P.Shape> shapes, P.PlaceholderShape targetPlaceholder)
    {
        foreach (var pShape in shapes)
        {
            var refPlaceholder = GetPlaceholderFromShape(pShape);
            if (AreTypeAndIndexEqual(refPlaceholder, targetPlaceholder))
            {
                return pShape;
            }
        }

        return null;
    }

    private static P.Shape? FindShapeByTypeOnly(IEnumerable<P.Shape> shapes, P.PlaceholderShape targetPlaceholder)
    {
        if (targetPlaceholder.Type?.Value is null)
        {
            return null;
        }

        return shapes.FirstOrDefault(shape =>
            GetPlaceholderFromShape(shape)?.Type?.Value == targetPlaceholder.Type.Value);
    }

    private static P.PlaceholderShape? GetPlaceholderFromShape(P.Shape shape)
    {
        return shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
            .GetFirstChild<P.PlaceholderShape>();
    }

    private static bool AreTypeAndIndexEqual(P.PlaceholderShape? placeholder1, P.PlaceholderShape? placeholder2)
    {
        if (placeholder1 == null || placeholder2 == null)
        {
            return false;
        }

        return placeholder1.Type?.Value == placeholder2.Type?.Value &&
               placeholder1.Index?.Value == placeholder2.Index?.Value;
    }
}