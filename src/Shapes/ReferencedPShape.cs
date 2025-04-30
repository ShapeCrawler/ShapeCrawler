using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal readonly ref struct ReferencedPShape(OpenXmlElement pShapeTreeElement)
{
    internal Transform2D ATransform2D()
    {
        var pShape = (P.Shape)pShapeTreeElement;
        var openXmlPart = pShape.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        if (openXmlPart is SlidePart sdkSlidePart)
        {
            var layoutPShape = LayoutPShapeOrNullOf(pShape, sdkSlidePart);
            if (layoutPShape != null && layoutPShape.ShapeProperties!.Transform2D != null)
            {
                return layoutPShape.ShapeProperties.Transform2D;
            }

            return this.MasterPShapeOf(pShape).ShapeProperties!.Transform2D!;
        }

        return this.MasterPShapeOf(pShape).ShapeProperties!.Transform2D!;
    }

    private static P.Shape? PShapeOrNullOf(IEnumerable<P.Shape> pShapes, P.PlaceholderShape source)
    {
        // Try to find a match based on specific conditions
        foreach (var pShape in pShapes)
        {
            var target = pShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>();
            
            // Skip shapes without placeholder information
            if (target == null)
            {
                continue;
            }

            // Check if the shape matches any of our matching conditions
            if (IsIndexMatch(source, target) || 
                IsBodyWithMatchingIndex(source, target) ||
                IsTitleMatch(source, target) ||
                IsCenteredTitleMatch(source, target) ||
                IsGeneralTypeMatch(source, target) ||
                IsTitleCenteredTitleMatch(source, target))
            {
                return pShape;
            }
        }

        // Fallback: try to find a shape with matching type
        return FindShapeByType(pShapes, source);
    }
    
    private static bool IsIndexMatch(P.PlaceholderShape source, P.PlaceholderShape target)
    {
        return source.Index is not null && 
               target.Index is not null && 
               source.Index == target.Index;
    }

    private static bool IsBodyWithMatchingIndex(P.PlaceholderShape source, P.PlaceholderShape target)
    {
        return source.Type?.Value == P.PlaceholderValues.Body &&
               source.Index is not null && 
               target.Index is not null &&
               source.Index == target.Index;
    }

    private static bool IsTitleMatch(P.PlaceholderShape source, P.PlaceholderShape target)
    {
        return source.Type?.Value == P.PlaceholderValues.Title && 
               target.Type! == P.PlaceholderValues.Title;
    }

    private static bool IsCenteredTitleMatch(P.PlaceholderShape source, P.PlaceholderShape target)
    {
        return source.Type?.Value == P.PlaceholderValues.CenteredTitle && 
               target.Type! == P.PlaceholderValues.CenteredTitle;
    }

    private static bool IsGeneralTypeMatch(P.PlaceholderShape source, P.PlaceholderShape target)
    {
        return source.Type != null && 
               target.Type != null && 
               source.Type.Equals(target.Type);
    }

    private static bool IsTitleCenteredTitleMatch(P.PlaceholderShape source, P.PlaceholderShape target)
    {
        return source.Type?.Value == P.PlaceholderValues.Title &&
               target.Type! == P.PlaceholderValues.CenteredTitle;
    }

    private static P.Shape? FindShapeByType(IEnumerable<P.Shape> pShapes, P.PlaceholderShape source)
    {
        return pShapes.FirstOrDefault(x =>
            x.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                .GetFirstChild<P.PlaceholderShape>()?.Type == source.Type);
    }

    private static P.Shape? LayoutPShapeOrNullOf(P.Shape pShape, SlidePart sdkSlidePart)
    {
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>();
        if (pPlaceholderShape == null)
        {
            return null;
        }

        var layoutPShapes =
            sdkSlidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!.Elements<P.Shape>();

        var referencedPShape = PShapeOrNullOf(layoutPShapes, pPlaceholderShape);
        if (referencedPShape != null)
        {
            return referencedPShape;
        }

        return null;
    }

    private P.Shape MasterPShapeOf(P.Shape pShape)
    {
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>() !;
        var openXmlPart = pShape.Ancestors<OpenXmlPartRootElement>().First().OpenXmlPart!;
        var masterPShapes = openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>(),
            _ => ((SlideLayoutPart)openXmlPart).SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>()
        };

        var referencedPShape = PShapeOrNullOf(masterPShapes, pPlaceholderShape);
        if (referencedPShape != null)
        {
            return referencedPShape;
        }

        // https://answers.microsoft.com/en-us/msoffice/forum/all/placeholder-master/0d51dcec-f982-4098-b6b6-94785304607a?page=3
        if (pPlaceholderShape.Index?.Value == 4294967295)
        {
            return masterPShapes.FirstOrDefault(x => x.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>()?.Index?.Value == 1) !;
        }

        throw new SCException("An error occurred while getting referenced master shape.");
    }
}