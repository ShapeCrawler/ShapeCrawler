using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using D = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Extensions
{
    internal static class CompositeElementExtensions
    {
        public static bool IsPlaceholder(this OpenXmlCompositeElement compositeElement)
        {
            return compositeElement.Descendants<P.PlaceholderShape>().Any();
        }

        /// <summary>
        ///     Gets non visual drawing properties values (cNvPr).
        /// </summary>
        /// <returns>(identifier, hidden, name)</returns>
        public static (int, bool, string) GetNvPrValues(this OpenXmlCompositeElement compositeElement)
        {
            // .First() is used instead .Single() because group shape can have more than one id for its child elements
            var cNvPr = compositeElement.GetNonVisualDrawingProperties();
            var id = (int) cNvPr.Id.Value;
            var name = cNvPr.Name.Value;
            var parsedHiddenValue = cNvPr.Hidden?.Value;
            var hidden = parsedHiddenValue != null && parsedHiddenValue == true;

            return (id, hidden, name);
        }

        public static P.NonVisualDrawingProperties GetNonVisualDrawingProperties(
            this OpenXmlCompositeElement compositeElement)
        {
            // Get <p:cNvSpPr>
            return compositeElement switch
            {
                P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties
                    .NonVisualDrawingProperties,
                P.Shape pShape => pShape.NonVisualShapeProperties.NonVisualDrawingProperties,
                P.Picture pPicture => pPicture.NonVisualPictureProperties.NonVisualDrawingProperties,
                P.GroupShape pGroupShape => pGroupShape.NonVisualGroupShapeProperties.NonVisualDrawingProperties,
                _ => throw new ShapeCrawlerException()
            };
        }

        /// <summary>
        ///     Gets <see cref="P.ApplicationNonVisualDrawingProperties" /> (p:nvPr).
        /// </summary>
        /// <param name="compositeElement"></param>
        /// <returns></returns>
        public static P.ApplicationNonVisualDrawingProperties ApplicationNonVisualDrawingProperties(
            this OpenXmlCompositeElement compositeElement)
        {
            return compositeElement switch
            {
                P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties
                    .ApplicationNonVisualDrawingProperties,
                P.Shape pShape => pShape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties,
                P.Picture pPicture => pPicture.NonVisualPictureProperties.ApplicationNonVisualDrawingProperties,
                _ => throw new ShapeCrawlerException()
            };
        }
    }
}