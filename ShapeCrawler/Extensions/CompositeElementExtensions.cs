using System.Linq;
using DocumentFormat.OpenXml;
using ShapeCrawler.Exceptions;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;

namespace ShapeCrawler.Extensions
{
    /// <summary>
    ///     Extension methods for <see cref="OpenXmlCompositeElement" /> instance.
    /// </summary>
    public static class CompositeElementExtensions
    {
        /// <summary>
        ///     Determines whether element is placeholder.
        /// </summary>
        /// <param name="compositeElement"></param>
        /// <returns></returns>
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
            var cNvPr = compositeElement.Descendants<P.NonVisualDrawingProperties>().First();
            var id = (int) cNvPr.Id.Value;
            var name = cNvPr.Name.Value;
            var parsedHiddenValue = cNvPr.Hidden?.Value;
            var hidden = parsedHiddenValue != null && parsedHiddenValue == true;

            return (id, hidden, name);
        }

        public static P.NonVisualDrawingProperties GetNonVisualDrawingProperties(
            this OpenXmlCompositeElement compositeElement)
        {
            return compositeElement switch
            {
                P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties
                    .NonVisualDrawingProperties,
                P.Shape pShape => pShape.NonVisualShapeProperties.NonVisualDrawingProperties,
                _ => throw new ShapeCrawlerException()
            };
        }

        public static P.ApplicationNonVisualDrawingProperties GetApplicationNonVisualDrawingProperties(
            this OpenXmlCompositeElement compositeElement)
        {
            return compositeElement switch
            {
                P.GraphicFrame pGraphicFrame => pGraphicFrame.NonVisualGraphicFrameProperties
                    .ApplicationNonVisualDrawingProperties,
                P.Shape pShape => pShape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties,
                _ => throw new ShapeCrawlerException()
            };
        }
    }
}