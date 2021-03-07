using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using ShapeCrawler.Extensions;
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
    public abstract class Shape //TODO: make it internal
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
            string customDataElement =
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

        public int Id => (int)PShapeTreeChild.GetNonVisualDrawingProperties().Id.Value;

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

        public virtual GeometryType GeometryType => GetGeometryType();

        private GeometryType GetGeometryType()
        {
            P.ShapeProperties spPr = PShapeTreeChild.Descendants<P.ShapeProperties>().First();// TODO: optimize
            A.Transform2D transform2D = spPr.Transform2D;
            if (transform2D != null)
            {
                A.PresetGeometry aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();

                // Placeholder can have transform on the slide, without having geometry
                if (aPresetGeometry == null && spPr.OfType<A.CustomGeometry>().Any())
                {
                    return GeometryType.Custom;
                }

                var name = aPresetGeometry.Preset.Value.ToString();
                Enum.TryParse(name, true, out GeometryType geometryType);
                return geometryType;
            }

            if (Placeholder != null)
            {
                return ((Placeholder)Placeholder).Shape.GeometryType;
            }

            return GeometryType.Rectangle; // return default
        }

        #endregion Public Properties
    }
}