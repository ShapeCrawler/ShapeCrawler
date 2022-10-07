using System;
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

        public SCPlaceholderType Type => this.GetPlaceholderType();
        
        protected internal Shape ReferencedShape => this.layoutReferencedShape.Value;

        #region Private Methods

        private SCPlaceholderType GetPlaceholderType()
        {
            var pPlaceholderValue = this.PPlaceholderShape.Type;
            if (pPlaceholderValue == null)
            {
                return SCPlaceholderType.Custom;
            }

            if (pPlaceholderValue == P.PlaceholderValues.Title)
            {
                return SCPlaceholderType.Title;
            }

            if (pPlaceholderValue == P.PlaceholderValues.CenteredTitle)
            {
                return SCPlaceholderType.CenteredTitle;
            }

            return (SCPlaceholderType)Enum.Parse(typeof(SCPlaceholderType), pPlaceholderValue.Value.ToString());
        }

        #endregion Private Methods
    }
}