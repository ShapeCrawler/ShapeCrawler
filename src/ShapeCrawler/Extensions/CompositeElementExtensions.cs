using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class CompositeElementExtensions
{
    internal static P.NonVisualDrawingProperties GetNonVisualDrawingProperties(
        this OpenXmlCompositeElement compositeElement)
    {
        // Get <p:cNvSpPr>
        return compositeElement switch
        {
            P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!,
            P.Shape pShape => pShape.NonVisualShapeProperties!.NonVisualDrawingProperties!,
            P.Picture pPicture => pPicture.NonVisualPictureProperties!.NonVisualDrawingProperties!,
            P.GroupShape pGroupShape => pGroupShape.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!,
            P.ConnectionShape pCxnSp => pCxnSp.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!,
            _ => throw new SCException()
        };
    }
    
    internal static P.NonVisualDrawingProperties GetNonVisualDrawingProperties(
        this OpenXmlElement xmlElement)
    {
        // Get <p:cNvSpPr>
        return xmlElement switch
        {
            P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!,
            P.Shape pShape => pShape.NonVisualShapeProperties!.NonVisualDrawingProperties!,
            P.Picture pPicture => pPicture.NonVisualPictureProperties!.NonVisualDrawingProperties!,
            P.GroupShape pGroupShape => pGroupShape.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!,
            P.ConnectionShape pCxnSp => pCxnSp.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!,
            _ => throw new SCException()
        };
    }

    internal static P.ApplicationNonVisualDrawingProperties GetPNvPr(this OpenXmlCompositeElement compositeElement)
    {
        return compositeElement switch
        {
            P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties!
                .ApplicationNonVisualDrawingProperties!,
            P.Shape pShape => pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!,
            P.Picture pPicture => pPicture.NonVisualPictureProperties!.ApplicationNonVisualDrawingProperties!,
            P.ConnectionShape pCxnSp => pCxnSp.NonVisualConnectionShapeProperties!.ApplicationNonVisualDrawingProperties!,
            P.GroupShape pGroupShape => pGroupShape.NonVisualGroupShapeProperties!.ApplicationNonVisualDrawingProperties!,
            _ => throw new SCException()
        };
    }
}