﻿using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Shapes;

internal readonly ref struct ReferencedPShape
{
    private readonly TypedOpenXmlPart sdkTypedOpenXmlPart;
    private readonly OpenXmlElement pShapeTreeElement;

    internal ReferencedPShape(TypedOpenXmlPart sdkTypedOpenXmlPart, OpenXmlElement pShapeTreeElement)
    {
        this.sdkTypedOpenXmlPart = sdkTypedOpenXmlPart;
        this.pShapeTreeElement = pShapeTreeElement;
    }

    internal Transform2D ATransform2D()
    {
        var pShape = (P.Shape)this.pShapeTreeElement;
        if (this.sdkTypedOpenXmlPart is SlidePart)
        {
            var layoutPShape = this.LayoutPShapeOrNullOf(pShape);
            if (layoutPShape != null)
            {
                return layoutPShape.ShapeProperties!.Transform2D!;
            }

            return this.MasterPShapeOf(pShape).ShapeProperties!.Transform2D!;
        }
        
        return this.MasterPShapeOf(pShape).ShapeProperties!.Transform2D!;
    }

    private P.Shape? LayoutPShapeOrNullOf(P.Shape pShape)
    {
        var pPlaceholderShape = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
            .GetFirstChild<P.PlaceholderShape>()!;
        var sdkSlidePart = (SlidePart)this.sdkTypedOpenXmlPart;
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
            .GetFirstChild<P.PlaceholderShape>()!;
        var masterPShapes = this.sdkTypedOpenXmlPart switch
        {
            SlidePart sdkSlidePart => sdkSlidePart.SlideLayoutPart!.SlideMasterPart!.SlideMaster.CommonSlideData!
                .ShapeTree!.Elements<P.Shape>(),
            _ => ((SlideLayoutPart)this.sdkTypedOpenXmlPart).SlideMasterPart!.SlideMaster.CommonSlideData!
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
                .GetFirstChild<P.PlaceholderShape>()?.Index?.Value == 1)!;
        }

        throw new Exception("An error occurred while getting referenced master shape.");
    }

    private static P.Shape? PShapeOrNullOf(IEnumerable<P.Shape> pShapes, P.PlaceholderShape pPlaceholderShape)
    {
        foreach (var pShape in pShapes)
        {
            var layoutPh = pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>();
            if (layoutPh == null)
            {
                continue;
            }

            if (pPlaceholderShape.Index is not null && layoutPh.Index is not null &&
                pPlaceholderShape.Index == layoutPh.Index)
            {
                return pShape;
            }

            if (pPlaceholderShape.Type == null || layoutPh.Type == null)
            {
                return pShape;
            }

            if (pPlaceholderShape.Type == P.PlaceholderValues.Body &&
                pPlaceholderShape.Index is not null && layoutPh.Index is not null)
            {
                if (pPlaceholderShape.Index == layoutPh.Index)
                {
                    return pShape;
                }
            }

            if (pPlaceholderShape.Type == P.PlaceholderValues.Title && layoutPh.Type == P.PlaceholderValues.Title)
            {
                return pShape;
            }
        }

        var byType = pShapes.FirstOrDefault(layoutPShape =>
            layoutPShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!
                .GetFirstChild<P.PlaceholderShape>()?.Type == pPlaceholderShape.Type);
        if (byType != null)
        {
            return byType;
        }

        return null;
    }
}