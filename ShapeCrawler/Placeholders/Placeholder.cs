using System;
using DocumentFormat.OpenXml;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler.Placeholders
{
    internal abstract class Placeholder : IPlaceholder
    {
        internal readonly P.PlaceholderShape PPlaceholderShape;

        protected Placeholder(P.PlaceholderShape pPlaceholderShape)
        {
            PPlaceholderShape = pPlaceholderShape;
        }

        protected internal Shape Shape { get; set; }

        public PlaceholderType Type => GetPlaceholderType();

        #region Private Methods

        private PlaceholderType GetPlaceholderType()
        {
            // Map SDK placeholder type into library placeholder type

            EnumValue<P.PlaceholderValues> pPlaceholderValue = PPlaceholderShape.Type;
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

        #endregion Private Methods
    }
}