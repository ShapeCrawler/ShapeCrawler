using System;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Groups;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Slides;

/// <summary>
///     Creates shape groupings.
/// </summary>
internal sealed class ShapeGrouping(SlidePart slidePart, ShapeIdGenerator idGenerator)
{
    /// <summary>
    ///     Creates a group from the specified shapes.
    /// </summary>
    internal IShape Create(IShape[] groupingShapes)
    {
        var nonVisualGroupShapeProperties = new P.NonVisualGroupShapeProperties();
        var idAndName = idGenerator.GenerateIdAndName();
        var nonVisualDrawingProperties = new P.NonVisualDrawingProperties
        {
            Id = (uint)idAndName.Item1, Name = idAndName.Item2
        };
        var nonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
        var applicationNonVisualDrawingProperties = new P.ApplicationNonVisualDrawingProperties();

        nonVisualGroupShapeProperties.Append(nonVisualDrawingProperties);
        nonVisualGroupShapeProperties.Append(nonVisualGroupShapeDrawingProperties);
        nonVisualGroupShapeProperties.Append(applicationNonVisualDrawingProperties);

        var groupShapeProperties = new P.GroupShapeProperties();

        decimal minX = decimal.MaxValue;
        decimal minY = decimal.MaxValue;
        decimal maxX = decimal.MinValue;
        decimal maxY = decimal.MinValue;

        foreach (var groupingShape in groupingShapes)
        {
            minX = Math.Min(minX, groupingShape.X);
            minY = Math.Min(minY, groupingShape.Y);
            maxX = Math.Max(maxX, groupingShape.X + groupingShape.Width);
            maxY = Math.Max(maxY, groupingShape.Y + groupingShape.Height);
        }

        var transformGroup = new A.TransformGroup();
        var offset = new A.Offset { X = (int)minX, Y = (int)minY };
        var extents = new A.Extents { Cx = (int)(maxX - minX), Cy = (int)(maxY - minY) };
        var childOffset = new A.ChildOffset { X = 0, Y = 0 };
        var childExtents = new A.ChildExtents { Cx = extents.Cx, Cy = extents.Cy };

        transformGroup.Append(offset);
        transformGroup.Append(extents);
        transformGroup.Append(childOffset);
        transformGroup.Append(childExtents);

        groupShapeProperties.Append(transformGroup);

        var pGroupShape = new P.GroupShape();
        pGroupShape.Append(nonVisualGroupShapeProperties);
        pGroupShape.Append(groupShapeProperties);

        foreach (var groupingShape in groupingShapes)
        {
            // Get the OpenXml element for the shape
            var openXmlElement = groupingShape.SDKOpenXmlElement;

            // Remove the shape from its current parent
            if (openXmlElement.Parent is not null)
            {
                openXmlElement.Remove();
            }
            
            pGroupShape.Append(openXmlElement);
        }

        slidePart.Slide.CommonSlideData!.ShapeTree!.Append(pGroupShape);

        foreach (var grouping in groupingShapes)
        {
            grouping.Remove();
        }

        return new GroupShape(pGroupShape);
    }
}
