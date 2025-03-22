using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions;

internal static class CompositeElementExtensions
{
    internal static P.NonVisualDrawingProperties NonVisualDrawingProperties(this OpenXmlElement openXmlElement)
    {
        return openXmlElement switch
        {
            P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties!.NonVisualDrawingProperties!,
            P.Shape pShape => pShape.NonVisualShapeProperties!.NonVisualDrawingProperties!,
            P.Picture pPicture => pPicture.NonVisualPictureProperties!.NonVisualDrawingProperties!,
            P.GroupShape pGroupShape => pGroupShape.NonVisualGroupShapeProperties!.NonVisualDrawingProperties!,
            P.ConnectionShape pCxnSp => pCxnSp.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties!,
            _ => throw new SCException()
        };
    }
}