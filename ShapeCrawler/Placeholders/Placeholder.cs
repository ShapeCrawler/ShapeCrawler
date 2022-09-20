using System;
using DocumentFormat.OpenXml;
using ShapeCrawler.Shared;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    internal abstract class Placeholder : IPlaceholder
    {
        internal readonly P.PlaceholderShape PPlaceholderShape;
        
        protected ResettableLazy<Shape> layoutReferencedShape;

        protected Placeholder(P.PlaceholderShape pPlaceholderShape)
        {
            this.PPlaceholderShape = pPlaceholderShape;
        }

        public PlaceholderType Type => this.GetPlaceholderType();

        /// <summary>
        ///     Gets referenced shape from lower level slide.
        /// </summary>
        protected internal Shape ReferencedShape => this.layoutReferencedShape.Value;

        #region Private Methods

        private PlaceholderType GetPlaceholderType()
        {
            var pPlaceholderValue = this.PPlaceholderShape.Type;
            if (pPlaceholderValue == null)
            {
                return PlaceholderType.Custom;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Title)
            {
                return PlaceholderType.Title;
            }

            if (pPlaceholderValue == P.PlaceholderValues.CenteredTitle)
            {
                return PlaceholderType.CenteredTitle;
            }

            // TODO: consider refactor the statement since it looks horrible
            return (PlaceholderType)Enum.Parse(typeof(PlaceholderType), pPlaceholderValue.Value.ToString());
        }

        #endregion Private Methods
    }
}