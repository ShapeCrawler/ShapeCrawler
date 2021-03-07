using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Placeholders;
using ShapeCrawler.SlideMaster;
using ShapeCrawler.Statics;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace ShapeCrawler
{
    /// <summary>
    ///     Represents a shape.
    /// </summary>
    public abstract class Shape
    {
        #region Constructors

        protected Shape(OpenXmlCompositeElement pShapeTreeChild)
        {
            PShapeTreeChild = pShapeTreeChild;
        }

        #endregion Constructors

        internal OpenXmlCompositeElement PShapeTreeChild { get; }

        protected void SetCustomData(string value)
        {
            var customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            PShapeTreeChild.InnerXml += customDataElement;
        }

        private string GetCustomData()
        {
            var pattern = @$"<{ConstantStrings.CustomDataElementName}>(.*)<\/{ConstantStrings.CustomDataElementName}>";
            var regex = new Regex(pattern);
            var elementText = regex.Match(PShapeTreeChild.InnerXml).Groups[1];
            if (elementText.Value.Length == 0)
            {
                return null;
            }

            return elementText.Value;
        }

        #region Public Properties

        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        /// <summary>
        ///     Gets placeholder. Returns <c>NULL</c> if the shape is not a placeholder.
        /// </summary>
        public abstract IPlaceholder Placeholder { get; }

        public abstract ThemePart ThemePart { get; }
        public abstract PresentationSc Presentation { get; }
        public abstract SlideMasterSc SlideMaster { get; }

        #endregion Public Properties
    }
}