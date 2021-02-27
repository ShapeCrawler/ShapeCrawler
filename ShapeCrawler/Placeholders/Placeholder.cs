using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Extensions;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    public class Placeholder
    {
        private readonly P.PlaceholderShape _pPlaceholderShape;
        private readonly OpenXmlCompositeElement _pShapeTreeChild;

        internal Placeholder(OpenXmlCompositeElement pShapeTreeChild, P.PlaceholderShape pPlaceholderShape)
        {
            _pShapeTreeChild = pShapeTreeChild;
            _pPlaceholderShape = pPlaceholderShape;
        }

        public PlaceholderType Type => GetPlaceholderType();

        private PlaceholderType GetPlaceholderType()
        {
            // Map SDK placeholder type into library placeholder type

            EnumValue<P.PlaceholderValues> pPlaceholderValue = _pPlaceholderShape.Type;
            if (pPlaceholderValue == null)
            {
                return PlaceholderType.Custom;
            }

            // Consider Title and Centered Title and Title as same
            if (pPlaceholderValue == P.PlaceholderValues.Title ||
                pPlaceholderValue == P.PlaceholderValues.CenteredTitle)
            {
                return PlaceholderType.Title;
            }

            //TODO: consider refactor the statement since it looks horrible
            return (PlaceholderType) Enum.Parse(typeof(PlaceholderType), pPlaceholderValue.Value.ToString());
        }

        internal bool TryGetFontSizeByParagraphLvl(int paragraphLvl, out int fontSize)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        ///     Create placeholder. Returns <c>NULL</c> if the specified shape is not placeholder.
        /// </summary>
        /// <param name="pShapeTreeChild"></param>
        /// <returns></returns>
        public static Placeholder Create(OpenXmlCompositeElement pShapeTreeChild)
        {
            P.PlaceholderShape pPlaceholderShape =
                pShapeTreeChild.GetApplicationNonVisualDrawingProperties().GetFirstChild<P.PlaceholderShape>();
            if (pPlaceholderShape == null)
            {
                return null;
            }

            return new Placeholder(pShapeTreeChild, pPlaceholderShape);
        }
    }
}