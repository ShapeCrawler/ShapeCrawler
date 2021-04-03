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
        internal OpenXmlCompositeElement PShapeTreeChild { get; }
        internal abstract ThemePart ThemePart { get; }


        #region Constructors

        protected Shape(OpenXmlCompositeElement pShapeTreeChild)
        {
            PShapeTreeChild = pShapeTreeChild;
        }

        #endregion Constructors

        public int Id => (int)PShapeTreeChild.GetNonVisualDrawingProperties().Id.Value;
        public string Name => PShapeTreeChild.GetNonVisualDrawingProperties().Name;
        public bool Hidden => DefineHidden();
        public string CustomData
        {
            get => GetCustomData();
            set => SetCustomData(value);
        }

        protected void SetCustomData(string value)
        {
            string customDataElement =
                $@"<{ConstantStrings.CustomDataElementName}>{value}</{ConstantStrings.CustomDataElementName}>";
            PShapeTreeChild.InnerXml += customDataElement;
        }

        #region Public Properties

        /// <summary>
        ///     Gets placeholder. Returns <c>NULL</c> if the shape is not a placeholder.
        /// </summary>
        public abstract IPlaceholder Placeholder { get; }

        public abstract SCPresentation Presentation { get; }
        public abstract SCSlideMaster SlideMaster { get; }

        public virtual GeometryType GeometryType => GetGeometryType();

        /// <summary>
        ///     Gets y-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long Y
        {
            get => GetY();
            set => SetY(value);
        }

        /// <summary>
        ///     Gets height of the shape.
        /// </summary>
        public long Height
        {
            get => GetHeight();
            set => SetHeight(value);
        }

        /// <summary>
        ///     Gets width of the shape.
        /// </summary>
        public long Width
        {
            get => GetWidth();
            set => SetWidth(value);
        }
        /// <summary>
        ///     Gets x-coordinate of the upper-left corner of the shape.
        /// </summary>
        public long X
        {
            get => GetX();
            set => SetX(value);
        }

        private bool DefineHidden()
        {
            bool? parsedHiddenValue = PShapeTreeChild.GetNonVisualDrawingProperties().Hidden?.Value;
            return parsedHiddenValue != null && parsedHiddenValue == true;
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

        private void SetX(long value)
        {
            A.Offset aOffset = PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                Shape placeholderShape = ((Placeholder) Placeholder).Shape;
                placeholderShape.X = value;
            }
            else
            {
                aOffset.X = value;
            }
        }

        private long GetX()
        {
            A.Offset aOffset = PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder) Placeholder).Shape.X;
            }

            return aOffset.X;
        }


        private void SetY(long value)
        {
            throw new NotImplementedException();
        }

        private long GetY()
        {
            A.Offset aOffset = PShapeTreeChild.Descendants<A.Offset>().FirstOrDefault();
            if (aOffset == null)
            {
                return ((Placeholder) Placeholder).Shape.Y;
            }

            return aOffset.Y;
        }

        private void SetWidth(long value)
        {
            throw new NotImplementedException();
        }

        private long GetWidth()
        {
            A.Extents aExtents = PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder) Placeholder).Shape.Width;
            }

            return aExtents.Cx;
        }

        private void SetHeight(long value)
        {
            throw new NotImplementedException();
        }

        private long GetHeight()
        {
            A.Extents aExtents = PShapeTreeChild.Descendants<A.Extents>().FirstOrDefault();
            if (aExtents == null)
            {
                return ((Placeholder) Placeholder).Shape.Height;
            }

            return aExtents.Cy;
        }

        private GeometryType GetGeometryType()
        {
            P.ShapeProperties spPr = PShapeTreeChild.Descendants<P.ShapeProperties>().First(); // TODO: optimize
            A.Transform2D transform2D = spPr.Transform2D;
            if (transform2D != null)
            {
                A.PresetGeometry aPresetGeometry = spPr.GetFirstChild<A.PresetGeometry>();

                // Placeholder can have transform on the slide, without having geometry
                if (aPresetGeometry == null)
                {
                    if (spPr.OfType<A.CustomGeometry>().Any())
                    {
                        return GeometryType.Custom;
                    }
                }
                else
                {
                    var name = aPresetGeometry.Preset.Value.ToString();
                    Enum.TryParse(name, true, out GeometryType geometryType);
                    return geometryType;
                }
            }

            Placeholder placeholder = (Placeholder) Placeholder;
            if (placeholder?.Shape != null)
            {
                return placeholder.Shape.GeometryType;
            }

            return GeometryType.Rectangle; // return default
        }

        #endregion Public Properties
    }
}