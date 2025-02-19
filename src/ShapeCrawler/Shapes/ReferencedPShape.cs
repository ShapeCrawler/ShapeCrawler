using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal readonly ref struct ReferencedPShape
{
    private readonly OpenXmlPart openXmlPart;
    private readonly OpenXmlElement pShapeTreeElement;

    internal ReferencedPShape(OpenXmlPart openXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.openXmlPart = openXmlPart;
        this.pShapeTreeElement = pShapeTreeElement;
    }

    internal Transform2D ATransform2D()
    {
        var pShape = (P.Shape)this.pShapeTreeElement;
        if (this.openXmlPart is SlidePart sdkSlidePart)
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
        foreach (var pShape in pShapes)
        {
            var target = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>();
            if (target == null)
            {
                continue;
            }

            if (source.Index is not null && target.Index is not null &&
                source.Index == target.Index)
            {
                return pShape;
            }

            if (source.Type == null || target.Type == null)
            {
                continue;
            }

            if (source.Type == P.PlaceholderValues.Body &&
                source.Index is not null && target.Index is not null)
            {
                if (source.Index == target.Index)
                {
                    return pShape;
                }
            }

            if (source.Type == P.PlaceholderValues.Title && target.Type == P.PlaceholderValues.Title)
            {
                return pShape;
            }

            if (source.Type == P.PlaceholderValues.CenteredTitle && target.Type == P.PlaceholderValues.CenteredTitle)
            {
                return pShape;
            }

            if (source.Type != null && target.Type != null && source.Type.Equals(target.Type))
            {
                return pShape;
            }

            if (source.Type != null && source.Type == P.PlaceholderValues.Title
                                   && target.Type != null && target.Type == P.PlaceholderValues.CenteredTitle)
            {
                return pShape;
            }
        }

        var byType = pShapes.FirstOrDefault(x =>
            x.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>()?.Type == source.Type);
        if (byType != null)
        {
            return byType;
        }

        return null;
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
        var masterPShapes = this.openXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>(),
            _ => ((SlideLayoutPart)this.openXmlPart).SlideMasterPart!.SlideMaster.CommonSlideData!
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