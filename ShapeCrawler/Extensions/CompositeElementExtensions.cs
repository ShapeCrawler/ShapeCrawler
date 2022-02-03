using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions
{
    internal static class CompositeElementExtensions
    {
        internal static bool IsPlaceholder(this OpenXmlCompositeElement compositeElement)
        {
            return compositeElement.Descendants<P.PlaceholderShape>().Any();
        }

        internal static P.NonVisualDrawingProperties GetNonVisualDrawingProperties(
            this OpenXmlCompositeElement compositeElement)
        {
            // Get <p:cNvSpPr>
            return compositeElement switch
            {
                P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties!
                    .NonVisualDrawingProperties,
                P.Shape pShape => pShape.NonVisualShapeProperties!.NonVisualDrawingProperties,
                P.Picture pPicture => pPicture.NonVisualPictureProperties!.NonVisualDrawingProperties,
                P.GroupShape pGroupShape => pGroupShape.NonVisualGroupShapeProperties.NonVisualDrawingProperties,
                P.ConnectionShape pCxnSp => pCxnSp.NonVisualConnectionShapeProperties!.NonVisualDrawingProperties,
                _ => throw new ShapeCrawlerException()
            };
        }

        /// <summary>
        ///     Gets <see cref="P.ApplicationNonVisualDrawingProperties" /> (p:nvPr).
        /// </summary>
        internal static P.ApplicationNonVisualDrawingProperties ApplicationNonVisualDrawingProperties(
            this OpenXmlCompositeElement compositeElement)
        {
            return compositeElement switch
            {
                P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties!
                    .ApplicationNonVisualDrawingProperties,
                P.Shape pShape => pShape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties,
                P.Picture pPicture => pPicture.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties,
                P.ConnectionShape pCxnSp => pCxnSp.NonVisualConnectionShapeProperties!.ApplicationNonVisualDrawingProperties, 
                _ => throw new ShapeCrawlerException()
            };
        }
    }
}